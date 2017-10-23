<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module명          : 회계 
'*  2. Function명        : 결의전표등록 
'*  3. Program ID        : a5101ma
'*  4. Program 이름      : 결의전표 등록 
'*  5. Program 설명      : 결의전표내역을 등록, 수정, 삭제, 조회 
'*  6. Comproxy 리스트   : 
'*  7. 최초 작성년월일   : 2000/09/22
'*  8. 최종 수정년월일   : 2001/02/12
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : 안혜진 
'* 11. 전체 comment      :
'*
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ag/AcctCtrl_ko441_1.vbs">							</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ag/AcctCtrl_ko441_2.vbs">							</SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">					</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
'⊙: 비지니스 로직 ASP명 
Const BIZ_PGM_ID      = "a5404mb1.asp"			'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'⊙: Grid Columns

'@Grid_Column
Dim C_ItemSeq	    																	'☆: Spread Sheet 의 Columns 인덱스 
Dim C_deptcd		
Dim C_deptPopup	
Dim C_deptnm	   	
Dim C_AcctCd		
Dim C_AcctPopup   
Dim C_AcctNm      
Dim C_DrCrFg      
Dim C_DrCrNm 
Dim C_DocCur     
Dim C_ExchRate	
Dim C_BalAmt		
Dim C_ItemAmt     
Dim C_ItemLocAmt  
Dim C_IsLAmtChange
Dim C_ItemDesc
Dim C_GL_DT	
Dim C_OpenGlNo	
Dim C_OpenGlItemSeq		
Dim C_MgntFg
Dim C_AcctCd2
Dim C_InternalCd
Dim C_Costcd
Dim C_OrgchangeId


'@Grid_Column vspdData2
Dim C_ItemSeq_2	    																	'☆: Spread Sheet 의 Columns 인덱스 
Dim C_deptcd_2		
Dim C_deptPopup_2	
Dim C_deptnm_2	   	
Dim C_AcctCd_2		
Dim C_AcctPopup_2   
Dim C_AcctNm_2      
Dim C_DrCrFg_2      
Dim C_DrCrNm_2 
Dim C_DocCur_2 
Dim C_DocCurPopup_2    
Dim C_ExchRate_2	
Dim C_BalAmt_2		
Dim C_ItemAmt_2     
Dim C_ItemLocAmt_2  
Dim C_IsLAmtChange_2
Dim C_ItemDesc_2	
Dim C_OpenGlNo_2	
Dim C_OpenGlItemSeq_2		
Dim C_MgntFg_2		
Dim C_AcctCd2_2
		
Const C_GLINPUTTYPE = "OD"

Const C_MENU_NEW_TAB1 = "1110000000011111"
Const C_MENU_NEW_TAB2 = "1110010000011111"
Const C_MENU_CRT_TAB1=	"1110100100111111"
Const C_MENU_CRT_TAB2=	"1110111100111111"
Const C_MENU_UPD_TAB1=	"1111000000011111"
Const C_MENU_UPD_TAB2=	"1111000000011111"

Const C_TAB1 = 1																		'☜: Tab의 위치 
Const C_TAB2 = 2

Const C_CONDFIELD = 0
Const C_VSPDDATA1 = 1
Const C_VSPDDATA2 = 2
Const C_VSPDDATA3 = 3
Const C_VSPDDATA4 = 4
Const C_VSPDDATA5 = 5
Const C_VSPDDATA6 = 6

Const C_POPUP_DEPT = 1
Const C_POPUP_ACCT = 2
Const C_POPUP_DOCCUR = 3
Const C_POPUP_DOCCUR2 = 4

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgCurrRow
Dim lgStrPrevKeyDtl
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgBlnExecDelete
Dim lgFormLoad
Dim lgQueryOk
Dim lgstartfnc

Dim intItemCnt		
'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= 
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim lgArrAcctForVat
Dim lgBlnGetAcctForVat
Dim lgCurrentTabFg
Dim lgPreToolBarTab1
Dim lgPreToolBarTab2
Dim lgIntMaxItemSeq

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgIntMaxItemSeq = 0
    'lgCurrentTabFg = C_TAB1  
    frm1.txtTempGlNo.focus 
End Sub


'========================================================================================================= 
Sub SetDefaultVal()


	' 현재 Page의 Form Element들을 Clear한다. 
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
    frm1.hCongFg.value = ""
    
    frm1.txttempGLDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat)
    frm1.hCongFg.value = "" 
    
    frm1.txtDocCur.value = Parent.gCurrency    
    frm1.cboGlType.value = "03"
        
    frm1.txtCommandMode.value = "CREATE"
    frm1.cboGlInputType.value = C_GLINPUTTYPE

	frm1.txtDeptCd.value	= Parent.gDepart

	frm1.hOrgChangeId.value = Parent.gChangeOrgId
	'현금계정을 가지고온다.
	Call GetCheckAcct
	
	frm1.txtTempGlNo.focus
	    	
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    Set gActiveElement = document.activeElement		
    
End Sub


'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "A", "COOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "COOKIE", "MA") %>
End Sub


'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case   UCase  (Trim(pvSpdNo))
	  Case   "A"
		C_ItemSeq			= 1																	'☆: Spread Sheet 의 Columns 인덱스 
		C_deptcd			= 2
		C_deptPopup			= 3
		C_deptnm	   		= 4
		C_AcctCd			= 5
		C_AcctPopup			= 6
		C_AcctNm			= 7
		C_DrCrFg			= 8
		C_DrCrNm			= 9
		C_DocCur			= 10
		C_ExchRate			= 11
		C_BalAmt			= 12
		C_ItemAmt			= 13
		C_ItemLocAmt		= 14
		C_IsLAmtChange		= 15
		C_ItemDesc			= 16
		C_GL_DT				= 17	
		C_OpenGlNo			= 18
		C_OpenGlItemSeq		= 19
		C_MgntFg			= 20
		C_AcctCd2			= 21
		C_internalcd		= 22
		C_costcd			= 23
		C_orgchangeid		= 24
								
	 Case   "B"
		C_ItemSeq_2			= 1																	'☆: Spread Sheet 의 Columns 인덱스 
		C_deptcd_2			= 2
		C_deptPopup_2		= 3
		C_deptnm_2	   		= 4
		C_AcctCd_2			= 5
		C_AcctPopup_2		= 6
		C_AcctNm_2			= 7
		C_DrCrFg_2			= 8
		C_DrCrNm_2			= 9
		C_DocCur_2			= 10
		C_DocCurPopup_2		= 11		
		C_ExchRate_2		= 12
		C_BalAmt_2			= 13
		C_ItemAmt_2			= 14
		C_ItemLocAmt_2		= 15
		C_IsLAmtChange_2	= 16
		C_ItemDesc_2		= 17
		C_OpenGlNo_2		= 18
		C_OpenGlItemSeq_2	= 19
		C_MgntFg_2			= 20
		C_AcctCd2_2			= 21
	End Select	
End Sub


'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call initSpreadPosVariables(pvSpdNo)

	With frm1
	
    Select Case   UCase  (Trim(pvSpdNo))
		Case   "A"
	
		.vspdData.MaxCols = C_OrgchangeId + 1
		.vspdData.Col = .vspdData.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.vspdData.ColHidden = True

		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20060205",,parent.gAllowDragDropSpread 
		ggoSpread.ClearSpreadData 

		.vspdData.ReDraw = False

		Call GetSpreadColumnPos(pvSpdNo)	

		'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)
		Call AppendNumberPlace("6","3","0")

        ggoSpread.SSSetFloat  C_ItemSeq,      " ", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd,       "부서코드"        , 10, , , 10
        ggoSpread.SSSetButton C_deptpopup
        ggoSpread.SSSetEdit   C_deptnm,       "부서명"          , 17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd,       "계정코드"        , 15, , , 18
		ggoSpread.SSSetButton C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,       "계정코드명"      , 20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg,       " "                   ,  8
	    ggoSpread.SSSetCombo  C_DrCrNm,       "차대구분"        , 10
	    ggoSpread.SSSetEdit   C_DocCur,       "거래통화"        , 10, , , 10, 2
		ggoSpread.SSSetFloat  C_ExchRate,     "환율"            , 15, Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat  C_BalAmt,		  "잔액"            , 15, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	    	    
	    ggoSpread.SSSetFloat  C_ItemAmt,      "금액"            , 15, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt,   "금액(자국)"      , 15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit   C_ItemDesc,     "적  요"          , 30, , , 128
		ggoSpread.SSSetEdit   C_IsLAmtChange, ""                    , 10	
		
		ggoSpread.SSSetDate   C_GL_DT,		  "발생일"          , 10, 2, parent.gDateFormat   	
    	ggoSpread.SSSetEdit   C_OpenGlNo,	  "미결전표번호"    , 10
		ggoSpread.SSSetEdit   C_OpenGlItemSeq,"미결전표항목번호", 10	'C_MgntFg
		ggoSpread.SSSetEdit   C_MgntFg,		  "미결계정여부"    , 10
		ggoSpread.SSSetDate   C_GL_DT,		  "발생일"          , 10, 2, parent.gDateFormat   
		ggoSpread.SSSetEdit   C_AcctCd2,      "계정코드2"       , 10
		ggoSpread.SSSetEdit  C_InternalCd,    "내부부서코드",     10
		ggoSpread.SSSetEdit  C_Costcd,		  "코스트센터",       10
		ggoSpread.SSSetEdit  C_OrgchangeId,   "조직변경ID",       10
		
		call ggoSpread.MakePairsColumn(C_deptcd,C_deptpopup)
		call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)
		
		Call ggoSpread.SSSetColHidden(C_ItemSeq,	C_ItemSeq,True)                          '☜: 차대구분 Hidden Column
		Call ggoSpread.SSSetColHidden(C_DrCrFg,		C_DrCrFg,True)                           '☜: 차대구분 Hidden Column
		Call ggoSpread.SSSetColHidden(C_IsLAmtChange,C_IsLAmtChange,True)                    '☜: 사용자가 Local 금액을 직접입력하였는지 
		Call ggoSpread.SSSetColHidden(C_AcctCd2,	C_AcctCd2,True)                          '☜: 차대구분 Hidden Column  
		Call ggoSpread.SSSetColHidden(C_OpenGlNo,	C_OpenGlNo,True)
		Call ggoSpread.SSSetColHidden(C_OpenGlItemSeq,C_OpenGlItemSeq,True)
		Call ggoSpread.SSSetColHidden(C_MgntFg,		C_MgntFg,True)
		Call ggoSpread.SSSetColHidden(C_GL_DT,C_GL_DT,True )
		Call ggoSpread.SSSetColHidden(C_OrgchangeId,C_OrgchangeId,True )
		Call ggoSpread.SSSetColHidden(C_InternalCd,C_InternalCd,True )
		Call ggoSpread.SSSetColHidden(C_Costcd,C_Costcd,True )
		.vspdData.ReDraw = True  
		Call SetSpreadLock("I", 0, 1, "")             
  
    
    Case   "B"

		.vspdData4.MaxCols = C_AcctCd2_2 + 1
		.vspdData4.Col = .vspdData4.MaxCols				'☜: 공통콘트롤 사용 Hidden Column
		.vspdData4.ColHidden = True

		ggoSpread.Source = .vspdData4
		ggoSpread.Spreadinit "V20060205",,parent.gAllowDragDropSpread 
		ggoSpread.ClearSpreadData 

		.vspdData4.ReDraw = False
		
		Call GetSpreadColumnPos(pvSpdNo)  
		
		

		'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)
		Call AppendNumberPlace("6","3","0")
		
        ggoSpread.SSSetFloat  C_ItemSeq_2,      " "                   , 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd_2,       "부서코드"        ,10, , , 10
        ggoSpread.SSSetButton C_deptpopup_2
        ggoSpread.SSSetEdit   C_deptnm_2,       "부서명"          ,17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd_2,       "계정코드"        ,15, , , 18
		ggoSpread.SSSetButton C_AcctPopup_2
		ggoSpread.SSSetEdit   C_AcctNm_2,       "계정코드명"      ,20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg_2,       " "                   , 8
	    ggoSpread.SSSetCombo  C_DrCrNm_2,       "차대구분"        ,10
		ggoSpread.SSSetEdit   C_DocCur_2,       "거래통화"        ,10, , , 10, 2
		ggoSpread.SSSetButton C_DocCurPopup_2
		ggoSpread.SSSetFloat  C_ExchRate_2,     "환율"            ,15, Parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat  C_BalAmt_2,		"잔액"            ,15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec	    	    
	    ggoSpread.SSSetFloat  C_ItemAmt_2,      "금액"            ,15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt_2,   "금액(자국)"      ,15, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit   C_ItemDesc_2,     "비  고"          ,30, , , 128
		ggoSpread.SSSetEdit   C_IsLAmtChange_2, ""                    ,10
		ggoSpread.SSSetEdit   C_OpenGlNo_2,		"미결전표번호"    ,10
		ggoSpread.SSSetEdit   C_OpenGlItemSeq_2,"미결전표항목번호",10
		ggoSpread.SSSetEdit   C_MgntFg_2,		"미결계정여부"    ,10
		ggoSpread.SSSetEdit   C_AcctCd2_2,      "계정코드2"       ,10
		
		
		
		call ggoSpread.MakePairsColumn(C_deptcd_2,C_deptpopup_2)
		call ggoSpread.MakePairsColumn(C_AcctCd_2,C_AcctPopup_2)
		Call ggoSpread.MakePairsColumn(C_DocCur_2,C_DocCurPopup_2)
		
		Call ggoSpread.SSSetColHidden(C_ItemSeq_2,		C_ItemSeq_2,True)                          '☜: 차대구분 Hidden Column
		Call ggoSpread.SSSetColHidden(C_DrCrFg_2,		C_DrCrFg_2,True)                           '☜: 차대구분 Hidden Column
		Call ggoSpread.SSSetColHidden(C_BalAmt_2,		C_BalAmt_2,True)
		Call ggoSpread.SSSetColHidden(C_IsLAmtChange_2,	C_IsLAmtChange_2,True)                    '☜: 사용자가 Local 금액을 직접입력하였는지 
		Call ggoSpread.SSSetColHidden(C_OpenGlNo_2,		C_OpenGlNo_2,True)
		Call ggoSpread.SSSetColHidden(C_OpenGlItemSeq_2,C_OpenGlItemSeq_2,True)
		Call ggoSpread.SSSetColHidden(C_MgntFg_2,		C_MgntFg_2,True)
		Call ggoSpread.SSSetColHidden(C_AcctCd2_2,		C_AcctCd2_2,True)                          '☜: 차대구분 Hidden Column  
		
		
		
		.vspdData4.ReDraw = True
		Call SetSpread4Lock ("I", 0, 1, "" )
    
	End Select 
   end with

End Sub


'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData
		Set objSpread = .vspdData
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
		Select Case   Index
			Case   0			
				ggoSpread.SpreadLock C_AcctNm,		lRow, C_AcctNm,		lRow2
				ggoSpread.SpreadLock C_AcctPopup,	lRow, C_AcctPopup,	lRow2
		        ggoSpread.SpreadLock C_AcctCd,		lRow, C_AcctCd,		lRow2
		        ggoSpread.SpreadLock C_deptnm,		lRow, C_deptnm,		lRow2
				ggoSpread.SSSetProtected		.vspdData.MaxCols,-1,-1		        
			Case   1
				ggoSpread.SpreadLock C_ItemSeq, lRow, C_AcctCd2, lRow2	'Item Grid 전체 Lock설정 
				ggoSpread.SSSetProtected		.vspdData.MaxCols,-1,-1				
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    
    End With
    
End Sub

Sub SetSpread4Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread
    
    With frm1
		ggoSpread.Source = .vspdData4
		Set objSpread = .vspdData4
		lRow2 = objSpread.MaxRows
		objSpread.Redraw = False
		Select Case   Index
			Case   0			
				ggoSpread.SpreadLock C_AcctNm,		lRow, C_AcctNm,		lRow2
				ggoSpread.SpreadLock C_AcctPopup,	lRow, C_AcctPopup,	lRow2
		        ggoSpread.SpreadLock C_AcctCd,		lRow, C_AcctCd,		lRow2
		        ggoSpread.SpreadLock C_deptnm,		lRow, C_deptnm,		lRow2
				ggoSpread.SSSetProtected		.vspdData4.MaxCols,-1,-1		        
			Case   1
				ggoSpread.SpreadLock C_ItemSeq, lRow, C_AcctCd2, lRow2	'Item Grid 전체 Lock설정 
				ggoSpread.SSSetProtected	.vspdData4.MaxCols,-1,-1				
		End Select

		objSpread.Redraw = True
		Set objSpread = Nothing
    
    End With
    
End Sub


'=======================================================================================================
Sub SetSpread2Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
	Dim objSpread
    
    With frm1
    ggoSpread.Source = .vspdData2
	Set objSpread = .vspdData2
	lRow2 = objSpread.MaxRows
	objSpread.Redraw = False
	
    Select Case    Index
		Case    0			
			ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1
		Case    1
			ggoSpread.SpreadLock 1, lRow, objSpread.MaxCols, lRow2	
			ggoSpread.SSSetProtected	.vspdData2.MaxCols,-1,-1
	End Select

    objSpread.Redraw = True
    Set objSpread = Nothing
    
    End With
End Sub


'=======================================================================================================
Sub SetSpread5Lock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
	Dim objSpread
    
    With frm1
    ggoSpread.Source = .vspdData5
	Set objSpread = .vspdData5
	lRow2 = objSpread.MaxRows
	objSpread.Redraw = False
	
    Select Case   Index
		Case   0
			ggoSpread.SSSetProtected	.vspdData5.MaxCols,-1,-1			
		Case   1
			ggoSpread.SpreadLock 1, lRow, objSpread.MaxCols, lRow2	
			ggoSpread.SSSetProtected	.vspdData5.MaxCols,-1,-1
	End Select

    objSpread.Redraw = True
    Set objSpread = Nothing
    
    End With
End Sub




'=======================================================================================================
Sub SetSpreadColor(Byval stsFg, Byval Index, ByVal pvStartRow,ByVal pvEndRow)
    With frm1

		if  pvEndRow = "" THEN	pvEndRow = pvStartRow
		
		
		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData		
		ggoSpread.SSSetProtected C_ItemSeq,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_deptcd,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_deptNm,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_AcctCd,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_AcctPopup,	pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_AcctNm,		pvStartRow	,pvEndRow   ' 계정코드명 
		ggoSpread.SSSetProtected C_DrCrNm,		pvStartRow	,pvEndRow
    	ggoSpread.SSSetProtected C_DocCur,		pvStartRow  ,pvEndRow
		ggoSpread.SSSetProtected C_BalAmt,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_ExchRate,	pvStartRow	,pvEndRow
		ggoSpread.SSSetRequired C_ItemAmt,		pvStartRow	,pvEndRow	' 금액 
		.vspdData.ReDraw = True
		
    End With

End Sub


Sub SetSpreadColor2(Byval stsFg, Byval Index, ByVal pvStartRow,ByVal pvEndRow)
    With frm1

		if  pvEndRow = "" THEN	pvEndRow = pvStartRow

		.vspdData4.ReDraw = False
		ggoSpread.Source = .vspdData4	
		ggoSpread.SSSetProtected	C_ItemSeq_2,	pvStartRow	,pvEndRow  ' 
		ggoSpread.SSSetProtected	C_deptNm_2,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected	C_AcctNm_2,		pvStartRow	,pvEndRow   ' 계정코드명	 
		ggoSpread.SSSetProtected	C_BalAmt_2,		pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected	C_ExchRate_2,	pvStartRow	,pvEndRow
		ggoSpread.SSSetRequired		C_Deptcd_2,		pvStartRow	,pvEndRow	   ' 부서코드		
		ggoSpread.SSSetRequired		C_AcctCd_2,		pvStartRow	,pvEndRow' 계정코드 
		ggoSpread.SSSetRequired		C_DrCrNm_2,		pvStartRow	,pvEndRow	' 차대구분 
'		ggoSpread.SSSetProtected	C_DocCur_2,		pvStartRow	,pvEndRow 
		ggoSpread.SSSetRequired		C_DocCur_2,		pvStartRow,	pvEndRow	   ' 통화						
		ggoSpread.SSSetRequired		C_ItemAmt_2,	pvStartRow	,pvEndRow	' 금액	
		.vspdData4.ReDraw = True
		
    End With

End Sub

'============================================================================================================
Sub InitComboBox()
	
	Err.clear
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboGlType ,lgF0  ,lgF1  ,Chr(11))
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
	 
End Sub


'============================================================================================================
Function InitComboBoxGrid(ByVal pvSpdNo)
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & "  order by minor_cd desc ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
	
	Select Case   pvSpdNo
	Case   "A"
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm
	Case   "B"
		ggoSpread.Source = frm1.vspdData4
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg_2
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm_2
	End Select
End Function
 
 
Function OpenPopUp(Byval pStrCode, Byval pIntPopUp, ByVal pIntVspdData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
	Dim iArrStrRet				'권한관리 추가   							  
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case   pIntPopUp
		
		Case   C_POPUP_ACCT													' Header명(1)			
			iArrParam(0) = "계정코드팝업"								' 팝업 명칭 
			iArrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			iArrParam(2) = pStrCode											' Code Condition
			iArrParam(3) = ""												' Name Cindition
			iArrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & "  "'미결계정은 반제계정으로 사용하지 못하게 함 AND ISNULL(A_ACCT.MGNT_FG,'') <> 'Y' "												' Where Condition
			iArrParam(5) = "계정코드"									' 조건필드의 라벨 명칭 

			iArrField(0) = "A_ACCT.Acct_CD"									' Field명(0)
			iArrField(1) = "A_ACCT.Acct_NM"									' Field명(1)
    		iArrField(2) = "A_ACCT_GP.GP_CD"								' Field명(2)
			iArrField(3) = "A_ACCT_GP.GP_NM"								' Field명(3)
			
			iArrHeader(0) = "계정코드"									' Header명(0)
			iArrHeader(1) = "계정코드명"									' Header명(1)
			iArrHeader(2) = "그룹코드"									' Header명(2)
			iArrHeader(3) = "그룹명"

		Case   C_POPUP_DOCCUR
			If frm1.txtDocCur.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
			iArrParam(0) = "통화코드 팝업"								' 팝업 명칭			
			iArrParam(1) = "B_Currency"	    								' TABLE 명칭 
			iArrParam(2) = pStrCode											' Code Condition
			iArrParam(3) = ""												' Name Cindition
			iArrParam(4) = ""												' Where Condition
			iArrParam(5) = "통화코드"									' 조건필드의 라벨 명칭 

			iArrField(0) = "Currency"	    								' Field명(0)
			iArrField(1) = "Currency_desc"	    							' Field명(1)
    
			iArrHeader(0) = "통화코드"									' Header명(0)
			iArrHeader(1) = "통화코드명"									' Header명(3)

		Case    C_POPUP_DOCCUR2

			iArrParam(0) = "통화코드 팝업"								' 팝업 명칭			
			iArrParam(1) = "B_Currency"	    								' TABLE 명칭 
			iArrParam(2) = pStrCode											' Code Condition
			iArrParam(3) = ""												' Name Cindition
			iArrParam(4) = ""												' Where Condition
			iArrParam(5) = "통화코드"									' 조건필드의 라벨 명칭 

			iArrField(0) = "Currency"	    								' Field명(0)
			iArrField(1) = "Currency_desc"	    							' Field명(1)
    
			iArrHeader(0) = "통화코드"									' Header명(0)
			iArrHeader(1) = "통화코드명"								' Header명(1)																				' Header명(1)

	End Select
    
	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
									 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
	  If pIntPopUp	= C_POPUP_DOCCUR Then
		frm1.txtDocCur.focus
	  End If	
		Exit Function
	Else
		Call SetPopUp(iArrRet, pIntPopUp, pIntVspdData)
	End If	

End Function


'--------------------------------------------------------------------------------------------------------- 

Function SetPopUp(Byval arrRet, Byval pIntPopUp, ByVal iVspdData)

	Dim iObjSpread
	
	If iVspdData = C_VSPDDATA1 Then
		Set iObjSpread = frm1.vspdData
	ElseIf iVspdData = C_VSPDDATA4 Then
		Set iObjSpread = frm1.vspdData4
	End If

	With frm1	
		Select Case   pIntPopUp				
			Case   C_POPUP_DOCCUR
				.txtDocCur.focus
				.txtDocCur.value = UCase  (Trim(arrRet(0)))				
				Call txtDocCur_OnChange()
			Case   C_POPUP_ACCT
				iObjSpread.Row = iObjSpread.ActiveRow 
			
				iObjSpread.Col  = C_AcctCD
				iObjSpread.Text = arrRet(0)
				iObjSpread.Col  = C_AcctNm
				iObjSpread.Text = arrRet(1)
				
				If iVspdData = C_VSPDDATA1 Then
					Call vspdData_Change(C_AcctCd, iObjSpread.Activerow)
				ElseIf iVspdData = C_VSPDDATA4 Then					
					Call vspdData4_Change(C_AcctCd_2, iObjSpread.Activerow)
				End If
			Case	C_POPUP_DOCCUR2
				iObjSpread.Row = iObjSpread.ActiveRow 
				iObjSpread.Col  = C_DocCur_2
				iObjSpread.Text = arrRet(0)
					Call vspdData4_Change(C_DocCur_2, iObjSpread.Activerow)
		End Select
	End With
	
End Function


'--------------------------------------------------------------------------------------------------------- 
Function OpenDept(Byval pStrCode)
	Dim iCalledAspName
	Dim iArrRet
	Dim iArrParam(8)

	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.className = parent.UCN_PROTECTED Then Exit Function

	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pStrCode									'  Code Condition
   	iArrParam(1) = frm1.txtTempGLDt.Text
	iArrParam(2) = lgUsrIntCd								' 자료권한 Condition  

	If lgIntFlgMode = Parent.OPMD_UMODE then
		iArrParam(3) = "T"									' 결의일자 상태 Condition  
	Else
		iArrParam(3) = "F"									' 결의일자 상태 Condition  
	End If

	' 권한관리 추가 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent, iArrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(iArrRet, C_CONDFIELD)
	End If	
			
End Function


Function OpenUnderDept(Byval pStrCode, ByVal pIntVspdData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
    Dim field_fg   	

	IsOpenPop = True
	If RTrim(LTrim(frm1.txtDeptCd.value)) <> "" 	Then
		iArrParam(0) = "부서 팝업"	
		iArrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"				
		iArrParam(2) = Trim(pStrCode)
		iArrParam(3) = "" 
		iArrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
		iArrParam(4) = iArrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD = ( SELECT B.BIZ_AREA_CD"
		iArrParam(4) = iArrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
		iArrParam(4) = iArrParam(4) & " AND A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"
		iArrParam(5) = "부서코드"			
	
		iArrField(0) = "A.DEPT_CD"	
		iArrField(1) = "A.DEPT_Nm"
		iArrField(2) = "B.BIZ_AREA_CD"
    
		iArrHeader(0) = "부서코드"
		iArrHeader(1) = "부서코드명"
		iArrHeader(2) = "사업장코드"
	Else
		iArrParam(0) = "부서 팝업"
		iArrParam(1) = "B_ACCT_DEPT A"				
		iArrParam(2) = Trim(pStrCode)
		iArrParam(3) = "" 
		iArrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		iArrParam(4) = iArrParam(4) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		iArrParam(4) = iArrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txttempGLDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			
		iArrParam(5) = "부서코드"
		iArrField(0) = "A.DEPT_CD"	
		iArrField(1) = "A.DEPT_Nm"
		iArrHeader(0) = "부서코드"
		iArrHeader(1) = "부서코드명"
	End IF

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(iArrRet, pIntVspdData)
	End If	
End Function


'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval pArrRet, ByVal pIntVspdData)
	
	Dim iObjSpread
	
	If pIntVspdData = C_VSPDDATA1 Then
		Set iObjSpread = frm1.vspdData
	ElseIf pIntVspdData = C_VSPDDATA4 Then
		Set iObjSpread = frm1.vspdData4
	End If
		
	With frm1
		If  pIntVspdData = C_CONDFIELD Then
			.txtDeptCd.focus
			.txtDeptCd.value = pArrRet(0)
			.txtDeptNm.value = pArrRet(1)
			.txtInternalCd.value = pArrRet(2)
  			 If lgQueryOk <> True Then
			 	.txtTempGLDt.text = pArrRet(3)
			 Else 
	
			 End If           
			 Call txtDeptCd_OnChange() 
		Else
			iObjSpread.Row = iObjSpread.ActiveRow 
				
			ggoSpread.Source = iObjSpread
			ggoSpread.UpdateRow iObjSpread.ActiveRow 
				
			iObjSpread.Col  = C_deptcd
			iObjSpread.Text = pArrRet(0)
			iObjSpread.Col  = C_deptnm
			iObjSpread.Text = pArrRet(1)
				
			Call deptCd_underChange("",pArrRet(0))
		End If				
	End With
	
End Function     


'=============================================================================================================== 
Function OpenRefTempGl()

	Dim iCalledAspName
	Dim iArrRet
	Dim iArrParam(8)	                           '권한관리 추가 (3 -> 4)

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a5404ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5404ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	iArrParam(4)	= lgAuthorityFlag              '권한관리 추가	
	
	' 권한관리 추가 
	iArrParam(5) = lgAuthBizAreaCd
	iArrParam(6) = lgInternalCd
	iArrParam(7) = lgSubInternalCd
	iArrParam(8) = lgAuthUsrID

	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If iArrRet(0) = ""  Then
		frm1.txtTempGlNo.focus			
		Exit Function
	Else		
		frm1.txtTempGlNo.focus			
		frm1.txttempGlNo.value = UCase(Trim(iArrRet(0)))   
	End If
	
End Function



'---------------------------------------------------------------------------------------------__------------
Function OpenPopupOA()
	
	Dim iArrRet	
	Dim iStrParm
	Dim iStrParm2
	Dim iStrParm3
	Dim arrParam(8)
	Dim ii
	Dim iCalledAspName

	If lgIntFlgMode = Parent.OPMD_UMODE Then Exit Function '@@수정추가시 지우기 

	iCalledAspName = AskPRAspName("a5403ra2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5403ra2", "X")
		IsOpenPop = False
		Exit Function
	End If

	With frm1		
		For ii = 1 To .vspdData.MaxRows
			.vspdData.Row = ii
			.vspdData.Col = C_OpenGlNo			
			iStrParm = iStrParm & .vspdData.Text & Parent.gColSep
			.vspdData.Col = C_OpenGlItemSeq
			iStrParm = iStrParm & .vspdData.Text & Parent.gRowSep
		Next
	End With

	iStrParm2 = frm1.txtDocCur.value
	iStrParm3 = frm1.txttempGLDt.text

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iStrParm,iStrParm2,iStrParm3,arrParam), _
		     "dialogWidth=900px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If iArrRet(0,0) = ""  Then			
		Exit Function
	Else		
		Call SetOpenPopupOA(iArrRet)
	End If

End Function

'--------------------------------------------------------------------------------------------------------------
Sub SetOpenPopupOA(ByVal arrRet)
	
	Dim ii
	Dim jj
	Dim tempstr
	Dim lngRows
	
	If lgCurrentTabFg = C_TAB2 Then
		Call ChangeTabs(C_TAB1)
		lgCurrentTabFg = C_TAB1
	End If
			
	With frm1
		ggoSpread.Source	= frm1.vspdData
		.vspdData.ReDraw	= False					
		For ii = 0 To Ubound(arrRet,1)
			If ii = 0 then
				.txtDocCur.value = 	arrRet(ii,10)	
				Call ggoOper.SetReqAttr(frm1.txtDocCur,	"Q")		
			End If						
			ggoSpread.InsertRow .vspdData.MaxRows
			lgIntMaxItemSeq = lgIntMaxItemSeq + 1			
			.vspdData.Row		= .vspdData.MaxRows

			For jj = 1 to C_OrgchangeId + 1
				.vspdData.Col = jj
				Select Case   Cstr(jj)
					Case   Cstr(C_ItemSeq)
						.vspdData.text = lgIntMaxItemSeq
					Case    Cstr(C_deptcd)
						.vspdData.text =  arrRet(ii,12)
					Case    Cstr(C_deptNm)
						.vspdData.text =  arrRet(ii,13)				
					Case   Cstr(C_AcctCd)
						.vspdData.text = arrRet(ii,2)						
					Case   Cstr(C_AcctNm)
						.vspdData.text = arrRet(ii,3)						

					Case   Cstr(C_DrCrFg)
						tempstr = arrRet(ii,4)
						If tempstr = "DR" Then			'DR
							.vspdData.value = 2
						ElseIf tempstr = "CR" Then		'CR
							.vspdData.value = 1
						End If
					Case   Cstr(C_DrCrNm)
						tempstr = arrRet(ii,4)  
						If tempstr = "DR" Then			'DR
							.vspdData.Value = 2
						ElseIf tempstr = "CR" Then		'CR
							.vspdData.Value = 1
						End If
					Case    Cstr(C_DocCur)
						.vspdData.text = arrRet(ii,10)
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.vspdData.Row,.vspdData.Row,C_DocCur, C_ExchRate,"D" ,"I","X","X")         		
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.vspdData.Row,.vspdData.Row,C_DocCur, C_BalAmt  ,"A" ,"I","X","X")         		
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,.vspdData.Row,.vspdData.Row,C_DocCur, C_ItemAmt ,"A" ,"I","X","X")      
					Case   Cstr(C_ExchRate)
						.vspdData.text = UNICDbl(arrRet(ii,8))
					Case   Cstr(C_OpenGlNo)
						.vspdData.text = arrRet(ii,0)						
					Case   Cstr(C_OpenGlItemSeq)	
						.vspdData.text = arrRet(ii,1)						
						Case   Cstr(C_MgntFg)	
						.vspdData.text = arrRet(ii,9)
					Case   Cstr(C_BalAmt)
						.vspdData.text = UNICDbl(arrRet(ii,6))						
					Case   Cstr(C_ItemAmt)
						.vspdData.text = UNICDbl(arrRet(ii,6))
					Case    Cstr(C_GL_DT)
						.vspdData.Text = arrRet(ii,11)
					Case   Cstr(C_ItemDesc)
						.vspdData.text = arrRet(ii,7)	
					Case    Cstr(C_InternalCd)
						.vspdData.text = arrRet(ii,14)
					Case    Cstr(C_Costcd)
						.vspdData.text = arrRet(ii,15)
					Case    Cstr(C_OrgchangeId)
						.vspdData.text = arrRet(ii,16)									
					Case   Else
						.vspdData.text = ""						
				End Select			
			Next

			Call vspdData_Change(C_ItemAmt,frm1.vspdData.ActiveRow)
			Call SetSpreadColor("I",0, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
		Next

		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DocCur, C_ExchRate,"D" ,"I","X","X")         		
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DocCur, C_BalAmt,"A" ,"I","X","X")         		
		Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_DocCur, C_ItemAmt,"A" ,"I","X","X")         		
		
		.vspdData.ReDraw = True
	End With
   
	For lngRows = 1 To frm1.vspdData.Maxrows    
		Call DbQuery2(lngRows, C_VSPDDATA1)
	Next 
    
    Call SetToolBar(C_MENU_CRT_TAB1)
    lgPreToolBarTab1 = C_MENU_CRT_TAB1
End Sub


'--------------------------------------------------------------------------------------------------------------
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(1)
	Dim IntRetCD
	Dim iCalledAspName

	
	iCalledAspName = AskPRAspName("a5120ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function
	
	arrParam(0) =  frm1.htxtTempGlNo.value	'전표번호 
	arrParam(1) = ""			      
	IsOpenPop = True
    
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function



'=======================================================================================================
Function ClickTab1()
	
	If lgCurrentTabFg = C_TAB1 Then Exit Function
	
	Call ChangeTabs(C_TAB1)	 '
	lgCurrentTabFg = C_TAB1
	
	If lgPreToolBarTab1 <> "" Then
		Call SetToolBar(lgPreToolBarTab1)
		Exit Function
	End If

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
	   Call SetToolBar(C_MENU_NEW_TAB1)
	Else				 
	End If
	
End Function


Function ClickTab2()
	
		
	If lgCurrentTabFg = C_TAB2 Then Exit Function
	Call ChangeTabs(C_TAB2)	 '~~~ 첫번째 Tab 
	lgCurrentTabFg = C_TAB2
	
	If lgPreToolBarTab2 <> "" Then
		Call SetToolBar(lgPreToolBarTab2)
		Exit Function
	End If
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		If frm1.vspdData4.MaxRows = 0 Then
			Call SetToolBar(C_MENU_NEW_TAB2)
		Else
			Call SetToolBar(C_MENU_CRT_TAB2)
		End If		
	ELSE				 
	END IF	

End Function
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'=======================================================================================================
'   Function Name : FindNumber
'   Function Desc : 
'=======================================================================================================
Function FindNumber(ByVal objSpread, ByVal intCol)
Dim lngRows
Dim lngPrevNum
Dim lngNextNum

    FindNumber = 0

    lngPrevNum = 0
    lngNextNum = 0
    
    With frm1
        
        If objSpread.MaxRows = 0 Then
            Exit Function
        End If
        
        For lngRows = 1 To objSpread.MaxRows
            objSpread.Row = lngRows
            objSpread.Col = intCol
            lngNextNum = Clng(objSpread.Text)
            
            If lngNextNum > lngPrevNum Then
                lngPrevNum = lngNextNum
            End If
            
        Next
       
    End With        
    
    FindNumber = lngPrevNum
    
End Function

'=======================================================================================================

Sub SetSpreadFG( pobjSpread , ByVal pMaxRows )
    Dim lngRows 
    
    For lngRows = 1 To pMaxRows
        pobjSpread.Col = 0
        pobjSpread.Row = lngRows
        pobjSpread.Text = ""
    Next
    
End Sub

'=======================================================================================================
Function SetSumItem()
    Dim DblTotDrAmt 
    DIm DblTotLocDrAmt
    Dim DblTotCrAmt 
    DIm DblTotLocCrAmt        
    Dim lngRows 
    
    DblTotDrAmt		= 0
	DblTotLocDrAmt	= 0
    DblTotCrAmt		= 0 
    DblTotLocCrAmt	= 0

	ggoSpread.Source = frm1.vspdData
	
    With frm1.vspdData 
	          
		If .MaxRows > 0 Then    
	        For lngRows = 1 To .MaxRows
	            .Row = lngRows
                    .Col = 0
                    if .text <> ggoSpread.DeleteFlag then

		            .col = C_DrCrFg
			    
		            if .text = "DR" then		
		
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If
			            			            
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
	
		            elseif .text = "CR" then
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If
			            
			            
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
		
			    end if	
		     end if	            
	        Next 
       End If                
	End With
	
	
	With frm1.vspdData4 
	          
		If .MaxRows > 0 Then    
	        For lngRows = 1 To .MaxRows
	            .Row = lngRows
                    .Col = 0
                    if .text <> ggoSpread.DeleteFlag then

		            .col = C_DrCrFg_2
			    
		            if .text = "DR" then		
		
			            .Col = C_ItemAmt_2	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If
			            			            
			            .Col = C_ItemLocAmt_2	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
	
		            elseif .text = "CR" then
			            .Col = C_ItemAmt_2	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If
			            
			            
			            .Col = C_ItemLocAmt_2	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
		
			    end if	
		     end if	            
	        Next 
       End If                
	End With
	
	
    frm1.txtDrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
    frm1.txtCrLocAmt.text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
    frm1.txtDrLocAmt2.text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
    frm1.txtCrLocAmt2.text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,Parent.gCurrency,Parent.ggAmtOfMoneyNo, Parent.gLocRndPolicyNo, "X")
      
	If frm1.cboGlType.value = "01" Then
		frm1.txtDrLocAmt.text	= frm1.txtCrLocAmt.text
		frm1.txtDrLocAmt2.text	= frm1.txtCrLocAmt2.text
	ElseIF frm1.cboGlType.value = "02" Then
		frm1.txtCrLocAmt.text	= frm1.txtDrLocAmt.text
		frm1.txtCrLocAmt2.text	= frm1.txtDrLocAmt2.text
	End If
	
End Function

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)

'	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim strTemp
	Dim strNmwhere
	Dim arrVal
	Dim IntRetCD
	
	Select Case   Kubun		
	Case   "FORM_LOAD"
	
		strTemp = ReadCookie("TEMP_GL_NO")
		Call WriteCookie("TEMP_GL_NO", "")

		If strTemp = "" then Exit Function
					
		frm1.txtTempGlNo.value = strTemp
				
		If Err.number <> 0 Then
			Err.Clear
			Call WriteCookie("TEMP_GL_NO", "")
			Exit Function 
		End If
				
		Call MainQuery()
	
	Case   Else
		Exit Function
	End Select
End Function


'========================================================================================================
'	Desc : 입출금 화면에 따른 Grid의 Protect변환 
'========================================================================================================
Sub CboGLType_ProtectGrid(Byval GlType)
	ggoSpread.Source = frm1.vspdData
	Select Case   GlType		
		Case   "01"			
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows	' 차대구분 
		Case   "02"			
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows	' 차대구분 
		Case   "03"			
			ggoSpread.SpreadUnLock C_DrCrfg, 1, C_DrCrNm, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrfg, 1, frm1.vspddata.maxrows	' 차대구분 
			ggoSpread.SSSetRequired C_DrCrNm, 1, frm1.vspddata.maxrows	' 차대구분 
	END Select 				
end Sub

'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case   UCase  (pvSpdNo)
		Case   "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemSeq			= iCurColumnPos(1)
			C_deptcd			= iCurColumnPos(2)
			C_deptPopup			= iCurColumnPos(3)
			C_deptnm			= iCurColumnPos(4)
			C_AcctCd			= iCurColumnPos(5)
			C_AcctPopup			= iCurColumnPos(6)
			C_AcctNm			= iCurColumnPos(7)
			C_DrCrFg			= iCurColumnPos(8)
			C_DrCrNm			= iCurColumnPos(9)
			C_DocCur            = iCurColumnPos(10) 
			C_ExchRate			= iCurColumnPos(11)
			C_BalAmt			= iCurColumnPos(12)
			C_ItemAmt			= iCurColumnPos(13)
			C_ItemLocAmt		= iCurColumnPos(14)
			C_IsLAmtChange		= iCurColumnPos(15)
			C_ItemDesc			= iCurColumnPos(16)
			C_GL_DT				= iCurColumnPos(17)
			C_OpenGlNo			= iCurColumnPos(18)
			C_OpenGlItemSeq		= iCurColumnPos(19)
			C_MgntFg			= iCurColumnPos(20)
			C_AcctCd2			= iCurColumnPos(21)
			C_internalcd		= iCurColumnPos(22)
			C_costcd			= iCurColumnPos(23)
			C_orgchangeid		= iCurColumnPos(24)
					
		 Case   "B"
            ggoSpread.Source = frm1.vspdData4
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemSeq_2			= iCurColumnPos(1)
			C_deptcd_2			= iCurColumnPos(2)
			C_deptPopup_2		= iCurColumnPos(3)
			C_deptnm_2			= iCurColumnPos(4)
			C_AcctCd_2			= iCurColumnPos(5)
			C_AcctPopup_2		= iCurColumnPos(6)
			C_AcctNm_2			= iCurColumnPos(7)
			C_DrCrFg_2			= iCurColumnPos(8)
			C_DrCrNm_2			= iCurColumnPos(9)
			C_DocCur_2		    = iCurColumnPos(10)
			C_DocCurPopup_2		= iCurColumnPos(11)  
			C_ExchRate_2		= iCurColumnPos(12)
			C_BalAmt_2			= iCurColumnPos(13)
			C_ItemAmt_2			= iCurColumnPos(14)
			C_ItemLocAmt_2		= iCurColumnPos(15)
			C_IsLAmtChange_2	= iCurColumnPos(16)
			C_ItemDesc_2		= iCurColumnPos(17)
			C_OpenGlNo_2		= iCurColumnPos(18)
			C_OpenGlItemSeq_2	= iCurColumnPos(19)
			C_MgntFg_2			= iCurColumnPos(20)
			C_AcctCd2_2			= iCurColumnPos(21)
			
    End Select    
End Sub

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'====================================================================================================
Sub  PopRestoreSpreadColumnInf()
	Dim indx

	ggoSpread.Source = gActiveSpdSheet
	Select Case   Trim(UCase  (gActiveSpdSheet.Name))
		Case   "VSPDDATA" 
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call InitComboBoxGrid("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData(C_VSPDDATA1)
			Call SetSpreadLock("Q", 1, 1, "")			
		Case   "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
'			Call SetSpread2Color()  
		Case   "VSPDDATA4" 
			Call PrevspdDataRestore2(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call InitComboBoxGrid("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData(C_VSPDDATA4)
			Call SetSpread4Lock("Q", 1, 1, "")
		Case   "VSPDDATA5"
			Call PrevspdData2Restore2(gActiveSpdSheet)   
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread2()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread4Color()  			
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If		
	
	If frm1.vspdData5.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData4.ActiveRow)
	End If			
End Sub

'====================================================================================================
Sub PrevspdDataRestore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case   frm1.vspdData.Text			
				Case   ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)
					Call PrevspdData2Restore(pActiveSheetName)
				Case   ggoSpread.UpdateFlag		
					For indx1 = 0 To frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case   frm1.vspdData3.Text 
							Case   ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase  (Trim(frm1.vspdData.Text)) = UCase  (Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)	 '''									
									Call FncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtTempGlNo.Value)
								End If
						End Select
					Next
				Case   ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtTempGlNo.Value)
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'====================================================================================================
Sub PrevspdData2Restore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case   frm1.vspdData2.Text
				Case   ggoSpread.InsertFlag
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
				Case   ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtTempGlNo.Value) 
						End If
					Next
				Case   ggoSpread.DeleteFlag
			End Select
		Else
			If lgIntFlgMode = Parent.OPMD_CMODE Then
				frm1.vspddata2.Maxrows = 0
			End If
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

									
'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ, "		
		strSelect = strSelect & " A.CTRL_CD, "
		strSelect = strSelect & " A.CTRL_NM, "
		strSelect = strSelect & " C.CTRL_VAL, "
		strSelect = strSelect & " '', "		
		strSelect = strSelect & " Case    WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END, "
		strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " Case  	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  END, " & strItemSeq & ","		
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
		strFrom	=			" A_CTRL_ITEM  		A (NOLOCK), "
		strFrom = strFrom & " A_ACCT_CTRL_ASSN  	B (NOLOCK), "
		strFrom = strFrom & " A_GL_DTL  			C (NOLOCK), "
		strFrom = strFrom & " A_GL_ITEM  			D (NOLOCK)	"
					
		strWhere =			  " D.GL_NO = " & FilterVar(UCase  (Trim(iStrOpenGlNo)), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ	= " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "		
		
				
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData(strItemSeq)  '''
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If
End Function

'====================================================================================================
Sub PrevspdDataRestore2(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData4.MaxRows
        frm1.vspdData4.Row = indx
        frm1.vspdData4.Col = 0
		
		If frm1.vspdData4.Text <> "" Then
			Select Case   frm1.vspdData4.Text			
				Case   ggoSpread.InsertFlag					
					frm1.vspdData4.Col = C_ItemSeq_2					
					Call DeleteHsheet2(frm1.vspdData4.Text)			'''		
				Case   ggoSpread.UpdateFlag		
					For indx1 = 0 To frm1.vspdData6.MaxRows					
						frm1.vspdData6.Row = indx1
						frm1.vspdData6.Col = 0
						Select Case   frm1.vspdData6.Text 
							Case   ggoSpread.UpdateFlag
								frm1.vspdData4.Col = C_ItemSeq_2
								frm1.vspdData6.Col = 1					
								If UCase  (Trim(frm1.vspdData4.Text)) = UCase  (Trim(frm1.vspdData6.Text)) Then
									Call DeleteHsheet2(frm1.vspdData4.Text)	 '''									
									Call FncRestoreDbQuery22(indx, frm1.vspdData4.ActiveRow, frm1.txtTempGlNo.Value)
								End If
						End Select
					Next
				Case   ggoSpread.DeleteFlag
					Call fncRestoreDbQuery22(indx, frm1.vspdData4.ActiveRow, frm1.txtTempGlNo.Value)
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'====================================================================================================
Sub PrevspdData2Restore2(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 to frm1.vspdData5.MaxRows
        frm1.vspdData5.Row    = indx
        frm1.vspdData5.Col    = 0

		If frm1.vspdData5.Text <> "" Then
			Select Case   frm1.vspdData5.Text
				Case   ggoSpread.InsertFlag
					frm1.vspdData5.Col = C_HItemSeq_2
					For indx1 = 0 to frm1.vspdData4.MaxRows
						frm1.vspdData4.Row = indx1
						frm1.vspdData4.Col = C_ItemSeq_2
						If frm1.vspdData4.Text = frm1.vspdData5.Text Then
							Call DeleteHsheet2(frm1.vspdData4.Text)
							ggoSpread.Source = frm1.vspdData4	
					        ggoSpread.EditUndo							
						End If
					Next
				Case   ggoSpread.UpdateFlag
					frm1.vspdData5.Col = C_HItemSeq_2
					For indx1 = 0 to frm1.vspdData4.MaxRows
						frm1.vspdData4.Row = indx1
						frm1.vspdData4.Col = C_ItemSeq_2
						If frm1.vspdData4.Text = frm1.vspdData5.Text Then
							Call DeleteHsheet2(frm1.vspdData4.Text)
							ggoSpread.Source = frm1.vspdData4
							ggoSpread.EditUndo
							Call fncRestoreDbQuery22(indx1, frm1.vspdData4.ActiveRow, frm1.txtTempGlNo.Value) 
						End If
					Next
				Case   ggoSpread.DeleteFlag
			End Select
		Else
			If lgIntFlgMode = Parent.OPMD_CMODE Then
				frm1.vspddata5.Maxrows = 0
			End If					
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub
							
'========================================================================================================
Function fncRestoreDbQuery22(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery22 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1
		.vspdData4.row = Row
	    .vspdData4.col = C_ItemSeq_2
		strItemSeq    = .vspdData4.Text

	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData4.Row = Row
		.vspdData4.Col = C_ItemSeq_2
		
		strSelect =				" C.DTL_SEQ, "		
		strSelect = strSelect & " A.CTRL_CD, "
		strSelect = strSelect & " A.CTRL_NM, "
		strSelect = strSelect & " C.CTRL_VAL, "
		strSelect = strSelect & " '', "		
		strSelect = strSelect & " Case    WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END, "
		strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
		strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " Case  	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  END, " & strItemSeq & ","		
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
		strFrom	=			" A_CTRL_ITEM  		A (NOLOCK), "
		strFrom = strFrom & " A_ACCT_CTRL_ASSN  	B (NOLOCK), "
		strFrom = strFrom & " A_GL_DTL  			C (NOLOCK), "
		strFrom = strFrom & " A_GL_ITEM  			D (NOLOCK)	"
					
		strWhere =			  " D.GL_NO = " & FilterVar(UCase  (Trim(iStrOpenGlNo)), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ	= " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "		
		
				
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData6
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData2(strItemSeq)  '''
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
	End With

	If Err.number = 0 Then
		fncRestoreDbQuery22 = True
	End If
End Function

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet("A")																'Setup the Spread sheet
    Call InitSpreadSheet("B")	
    Call InitCtrlSpread()
	Call InitCtrlHSpread()
	Call InitCtrlSpread2()
	Call InitCtrlHSpread2()															'Setup the Spread sheet  
        '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call InitComboBoxGrid("A")     
	Call InitComboBoxGrid("B")     
    Call SetAuthorityFlag                                      					
    Call SetDefaultVal    
    Call InitVariables 
    Call FncNew()
	Call CookiePage("FORM_LOAD")  	
	
	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing
    
End Sub


'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'=======================================================================================================
Sub vspdData_onfocus()

End Sub


'=======================================================================================================
Sub vspdData4_onfocus()

End Sub


'=======================================================================================================
Sub txttempGLDt_DblClick(Button)
    If Button = 1 Then
        frm1.txttempGLDt.Action = 7
    End If
End Sub


'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData 
    
    if Row = 0 or Col <> C_AcctCd then
	    Exit Sub
	End if
	
	ggoSpread.Source = frm1.vspdData
	frm1.vspddata.row = frm1.vspddata.ActiveRow	

 	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) < 1 Then
		frm1.vspdData2.Maxrows = 0
	end if
	
End Sub


'=======================================================================================================
Sub vspdData4_Click(ByVal Col, ByVal Row)
	
	Select Case   lgIntFlgMode
	Case   Parent.OPMD_UMODE
		Call SetPopUpMenuItemInf("0000111111")
	Case   Parent.OPMD_CMODE
		Call SetPopUpMenuItemInf("1101111111")
	End Select
	
    gMouseClickStatus = "SP2C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData4 

	If frm1.vspdData4.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData4
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub

    End If
	ggoSpread.Source = frm1.vspdData4
	frm1.vspddata4.row = frm1.vspddata4.ActiveRow	

 	frm1.vspdData4.Col = C_AcctCd_2
	
    If Len(frm1.vspdData4.Text) < 1 Then
		frm1.vspdData5.Maxrows = 0
	end if
	
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


'==========================================================================================
Sub vspdData4_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'========================================================================================== 
' Event Name : vspdData_LeaveCell 
' Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'			   Item Row 변경시 관리항목 처리 
'			   hItemSeq에 Item Seq 입력 
'			   lgCurrRow에 Row Index 입력 
'==========================================================================================

Sub vspdData_scriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

     If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq
            .hItemSeq.value = .vspdData.Text
            .vspdData2.MaxRows = 0
        End With

        frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End if
        		
		lgCurrRow = NewRow     		
		Call DbQuery2(lgCurrRow, C_VSPDDATA1)
		
    End If
End Sub

'==========================================================================================

Sub vspdData4_scriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData4.Row = NewRow
            .vspdData4.Col = C_ItemSeq_2
            .hItemSeq.value = .vspdData4.Text
            .vspdData5.MaxRows = 0
        End With

        frm1.vspddata4.Col = 0
        If frm1.vspddata4.Text = ggoSpread.DeleteFlag Then
			Exit Sub
		End if
        		
		lgCurrRow = NewRow     		
		Call DbQuery2(lgCurrRow, C_VSPDDATA4)
    End If
End Sub


'==========================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	'---------- Coding part -------------------------------------------------------------
	With frm1
		If Row > 0 And Col = C_AcctPopUp Then
			.vspdData.Col = Col - 1
			.vspdData.Row = Row

			Call OpenPopUp(.vspdData.text, C_POPUP_ACCT, C_VSPDDATA1)
		End If
		
		If Row > 0 And Col = C_deptPopup Then
			.vspdData.Col = Col - 1
			.vspdData.Row = Row							
			Call OpenUnderDept(.vspdData.Text, C_VSPDDATA1)
    	End If 
    	Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	   	
	End With
End Sub


'==========================================================================================

Sub vspdData4_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	'---------- Coding part -------------------------------------------------------------
	
	With frm1
	
		If Row > 0 And Col = C_AcctPopUp Then
			.vspdData4.Col = Col - 1
			.vspdData4.Row = Row
			
			Call OpenPopUp(.vspdData4.text, C_POPUP_ACCT, C_VSPDDATA4)
		End If
		
		If Row > 0 And Col = C_DocCurPopup_2 Then
			.vspdData4.Col = Col - 1
			.vspdData4.Row = Row			
			Call OpenPopUp(.vspdData4.text, C_POPUP_DOCCUR2, C_VSPDDATA4)
		End If
		
		If Row > 0 And Col = C_deptPopup Then
			.vspdData4.Col = Col - 1
			.vspdData4.Row = Row							
			Call OpenUnderDept(.vspdData4.Text, C_VSPDDATA4)
    	End If
    	Call SetActiveCell(.vspdData4,Col-1,.vspdData4.ActiveRow ,"M","X","X")   	    	
	End With
End Sub


'=======================================================================================================
Sub vspdData_Change(ByVal pCol, ByVal pRow)

	Dim IntRetCD
	Dim DeptCD	

	With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.UpdateRow pRow    
		.vspdData.Row = pRow   
    
		Select Case   pCol
		    Case     C_DeptCd
				.vspdData.Col = C_DeptCd
				DeptCD			= .vspdData.Text
				If DeptCd <> "" Then
					Call DeptCd_underChange("0",.vspdData.text)
				End If
		    Case     C_AcctCd
			    .vspdData.Col = 0
				If  .vspdData.Text = ggoSpread.InsertFlag Then
				
					.vspdData.Col = C_ItemSeq
					frm1.hItemSeq.value = .vspdData.Text
					.vspdData.Col = C_AcctCd							
										
					If Len(.vspdData.Text) > 0 Then
			
						.vspdData.Row = pRow					
						.vspdData.Col = C_ItemSeq	 						
						Call DeleteHsheet(.vspdData.Text)
	
					Else
						.vspdData.Col = C_AcctNm
						.vspdData.Text = ""
					End If   
			         
				End If
				
		  	Case  	C_DrCrFg
				Call SetSumItem()
			Case  	C_DrCrNm  
				Call SetSumItem()
			Case     C_ItemAmt	
				Call SetSumItem()	
			Case     C_ItemLocAmt
				.vspdData.Row = pRow
				.vspdData.Col = C_IsLAmtChange
				.vspdData.Text = "Y"
				Call SetSumItem()	
		End Select
    End WiTh	  
End Sub


'=======================================================================================================
Sub vspdData4_Change(ByVal pCol, ByVal pRow)	
	Dim tmpAcctCd
	Dim IntRetCD
	Dim iObjSpread
	Dim tmpDrCrFg
	Dim CurrencyCode
	Dim DeptCd   
	
	With frm1
		ggoSpread.Source = .vspdData4    
		ggoSpread.UpdateRow pRow    
		.vspdData4.Row = pRow   
    
		Select Case   pCol
		    Case     C_DeptCd_2
				.vspdData4.Col = C_DeptCd_2
				DeptCd         = .vspdData4.Text  
				If DeptCd <> "" Then
					Call DeptCd_underChange("1", .vspdData4.text)
				End If
		    Case    C_AcctCd_2
			    .vspdData4.Col = 0
				If  .vspdData4.Text = ggoSpread.InsertFlag Then
					.vspdData4.Col = C_ItemSeq_2
					frm1.hItemSeq.value = .vspdData4.Text
					.vspdData4.Col = C_AcctCd_2							
										
					If Len(.vspdData4.Text) > 0 Then
						.vspdData4.Row = pRow					
						.vspdData4.Col = C_ItemSeq_2	 
						Call DeleteHsheet2(.vspdData4.Text)
						
						.vspdData4.Row = pRow	
						.vspdData4.Col = C_DrCrFg_2		
						tmpDrCrFG = .vspdData4.text
						.vspdData4.Row = pRow
						.vspdData4.Col = C_AcctCd_2
						
						If AcctCheck2(.vspdData4.text, frm1.cboGlType.value, tmpDrCrFG) = True Then					
							Call Dbquery4(pRow)
							Call InputCtrlVal(pRow, C_VSPDDATA4)
							Call SetSpread4Color()
						End If
					Else
						.vspdData4.Col = C_AcctNm_2
						.vspdData4.Text = ""
					End If   
				End If
		  	Case 	C_DrCrFg_2
				Call SetSumItem()
			Case  	C_DrCrNm_2  
				Call SetSumItem()
			Case	C_DocCur_2	
				.vspdData4.Col = C_DocCur_2
				CurrencyCode = .vspdData4.Text
				If CurrencyCode <> "" Then
					IntRetCD = CommonQueryRs("Currency","B_CURRENCY"," Currency =  " & FilterVar(CurrencyCode , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					If IntRetCD = False Then
						Call DisplayMsgBox("am0028","X","X","X")
						.vspdData4.Col	= C_DocCur_2
						.vspdData4.Text = ""
					Else
						.vspdData4.Col = C_ItemLocAmt_2
						.vspdData4.Text = ""
						.txtDrLocAmt2.text = ""
						.vspdData4.Col= C_DocCur_2
						Call DocCur_OnChange(frm1.vspdData4.text, pRow )
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,pRow,pRow,C_DocCur_2,C_ExchRate_2,"D","I","X","X")  
						Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData4,pRow,pRow,C_DocCur_2,C_ItemAmt_2 ,"A","I","X","X")  
					End If
				End If
			Case  C_ExchRate_2	
				Call FixDecimalPlaceByCurrency(frm1.vspdData4,Row,C_DocCur_2,C_ItemAmt_2,  "A" ,"X","X")
			Case  C_ItemAmt_2	
				Call SetSumItem()	
			Case  C_ItemLocAmt_2
				.vspdData4.Row = pRow
				.vspdData4.Col = C_IsLAmtChange_2
				.vspdData4.Text = "Y"
				Call SetSumItem()	
		End Select
    End With
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	
	With frm1
		.vspddata.Row = Row		
		Select Case   Col
			Case   C_DrCrNm
				.vspddata.Col = Col				
				intIndex = .vspddata.Value
				.vspddata.Col = C_DrCrFg
				.vspddata.Value = intIndex
							
'				SetSpread2Color 
			Case   C_DrCrFg								
				.vspddata.Col = Col
				intIndex = .vspddata.Value				
				.vspddata.Col = C_DrCrNm
				.vspddata.Value = intIndex
		End Select		
	End With

End Sub


'==========================================================================================
Sub vspdData4_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim tmpDrCrFg
		  
	With frm1
		.vspddata4.Row = Row
		
		Select Case   Col
			Case   C_DrCrNm_2
				.vspddata4.Col = Col
				
				intIndex = .vspddata4.Value
				.vspddata4.Col = C_DrCrFg_2
				.vspddata4.Value = intIndex
				tmpDrCrFg = .vspddata4.text
				
				.vspddata4.Col = C_AcctCd_2
'				IF AcctCheck2(frm1.vspdData4.text,frm1.cboGlType.value, tmpDrCrFg) = True Then					
'					Call SetSpread2Color 					
'				END IF
'				SetSpread2Color 
			Case   C_DrCrFg_2
				.vspddata4.Col = Col
				
				intIndex = .vspddata4.Value
				.vspddata4.Col = C_DrCrNm_2
				.vspddata4.Value = intIndex
		End Select		
	End With

End Sub


'==========================================================================================
Sub txtTempGlNo_OnKeyPress()	
	If window.event.keycode = 39 then	'Single quotation mark 입력불가 
		window.event.keycode = 0	
	End If
End Sub


'==========================================================================================
Sub txtTempGlNo_OnKeyUp()	
	If Instr(1,frm1.txtTempGlNo.value,"'") > 0 then
		frm1.txtTempGlNo.value = Replace(frm1.txtTempGlNo.value, "'", "")		
	End if
End Sub


'==========================================================================================
Sub txtTempGlNo_onpaste()	
	Dim iStrTempGlNo 	
	iStrTempGlNo = window.clipboardData.getData("Text")
	iStrTempGlNo = RePlace(iStrTempGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrTempGlNo)		
End Sub


'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(frm1.txtDocCur.value , "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
'		Call CurFormatNumSprSheet()
		Call SetSumItem
	END IF	    
End Sub

'==========================================================================================

Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtTempGLDt.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then      
		Exit sub
    End If
    
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			

		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
	
		'----------------------------------------------------------------------------------------

End Sub

'==========================================================================================

Sub DeptCd_underChange(Byval stsfg, Byval strCode)
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 

    If Trim(frm1.txtTempGLDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtTempGLDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		
		select Case stsfg
			Case "0" 
			frm1.vspdData.Col = C_deptcd			
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.text = ""
			frm1.vspdData.Col = C_deptnm		
			frm1.vspdData.Row = frm1.vspdData.ActiveRow	
			frm1.vspdData.text = ""
			Case "1"
		    frm1.vspdData4.Col = C_deptcd_2			
			frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
			frm1.vspdData4.text = ""
			frm1.vspdData4.Col = C_deptnm_2		
			frm1.vspdData4.Row = frm1.vspdData4.ActiveRow	
			frm1.vspdData4.text = ""
		End Select	
	End If 
	
	'----------------------------------------------------------------------------------------

End Sub



'==========================================================================================

Sub txttempGLDt_Change()

   If lgstartfnc = False Then
    If lgFormLoad = True Then
		Dim strSelect
		Dim strFrom
		Dim strWhere 	
		Dim IntRetCD 
		Dim ii
		Dim arrVal1
		Dim arrVal2
		Dim jj


		lgBlnFlgChgValue = True
		With frm1
		
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtTempGLDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
			strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
			strFrom		=			 " b_acct_dept(NOLOCK) "		
			strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
			strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtTempGLDt.Text, Parent.gDateFormat,""), "''", "S") & "))"			
	
				If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
					IntRetCD = DisplayMsgBox("124600","X","X","X")
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					If .vspdData.MaxRows <> 0 Then
						For ii = 1 To .vspdData.MaxRows
						.vspdData.Col = C_deptcd			
					    .vspdData.Row = ii
					    .vspdData.text = ""
					    .vspdData.Col = C_deptnm	
					    .vspdData.text = ""
						Next		
					End If
				Else
					arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
					jj = Ubound(arrVal1,1)
							
					For ii = 0 to jj - 1
						arrVal2 = Split(arrVal1(ii), chr(11))			
						frm1.hOrgChangeId.value = Trim(arrVal2(2))
					Next	
				End If 
			End If
		End With
		'----------------------------------------------------------------------------------------
	End If
  End IF
  ggoSpread.Source = frm1.vspdData4	
  frm1.vspdData4.Col = C_DocCur_2
  If frm1.vspdData4.value <> "" Then 
  Call vspdData4_Change(C_DocCur_2, frm1.vspddata4.activerow)    
  End If
  
End Sub


'==========================================================================================
Sub cboGLType_OnChange()
	
	dim	i		
	Dim IntRetCD	
	
	ggoSpread.Source = frm1.vspdData
	
	SELECT Case   frm1.cboGlType.value 
		Case   "01"			
			'입금전표로 바꾸면 차변이 입력되거나 현금계정이 입력되었는지 check한다.
			FOR i = 1 TO  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				IF  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End IF
																			
				frm1.vspddata.col = C_DrCrFg
				IF  Trim(frm1.vspddata.value) = "2" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113104", "X", "X", "X")					
					Exit sub
				End IF											
			Next				
			
			FOR i = 1 TO  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				IF Trim(frm1.vspddata.value) <> "1"  Then					
					frm1.vspdData.value	= "1"							
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.value	= "1"							
				END IF
			Next
			
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		Case   "02"
			'출금전표로 바꾸면 대변이 입력되거나 현금계정이 입력되었는지 check한다.	
			FOR i = 1 TO  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				IF  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End IF								
				
				frm1.vspddata.col = C_DrCrFg
				IF  Trim(frm1.vspddata.value) = "1" Then
					frm1.cboGlType.value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
					IntRetCD = DisplayMsgBox("113105", "X", "X", "X")					
					Exit sub				
				End IF											
			Next
				
			FOR i = 1 TO  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				IF Trim(frm1.vspddata.value) <> "2"  Then					
					frm1.vspdData.value	= "2"							
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.value	= "2"							
				END IF
			Next
			
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		Case   "03"
		'대체로 바꾸면 Protect를 풀어준다.		
			Call CboGLType_ProtectGrid(frm1.cboGlType.value )		
		
	END SELECT	
	
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub


'========================================================================================================
Sub vspdData4_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub
  


'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If   
    	
  
    
End Sub


'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    Dim RetFlag

    lgstartfnc = True
    FncQuery = False          '⊙: Processing is NG
    Err.Clear                 '☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
 
     ' 변경된 내용이 있는지 확인한다.
    ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then
      		Exit Function
    	End If
    End If
 	
	' Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
	' ChkField(pDoc, pStrGrp) As Boolean
    If Not chkField(Document, "1") Then	'⊙: This function check indispensable field
       Exit Function
    End If
   '-----------------------
    'Erase contents area
    '-----------------------
	' 현재 Page의 Form Element들을 Clear한다. 

    Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear
 
    Call InitVariables							'⊙: Initializes local global variables

      
    '-----------------------
    'Query function call area
    '-----------------------
    IF  DbQuery = False Then														'☜: Query db data
		Exit Function
	END IF
		
    if frm1.vspddata.maxrows = 0 then	
       frm1.txtTempGlNo.value = ""
    end if
   
    FncQuery = True																'⊙: Processing is OK
    lgstartfnc = False
End Function

'========================================================================================
Function FncNew() 
	Dim IntRetCD 

	Dim var1, var2
	    
    FncNew = False                                                          
   lgstartfnc = True
    Err.Clear                       '☜: Protect system from crashing
    On Error Resume Next            '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    
	'-----------------------
    'Check previous data area
    '----------------------- 
    If (lgBlnFlgChgValue = True Or var1 = True Or var2 = True) And lgBlnExecDelete <> True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	lgBlnExecDelete = False
	
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
	
    Call ggoOper.ClearField(Document, "1")                                  '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '⊙: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    
	Call InitComboBoxGrid("A")     
	Call InitComboBoxGrid("B")     
    Call InitComboBoxGridVat    
    frm1.txtTempGlNo.focus        
    
    Call ClickTab1()
    Call SetToolbar(C_MENU_NEW_TAB1)
    lgPreToolBarTab1 = 	C_MENU_NEW_TAB1		 		    
    lgPreToolBarTab2 = 	C_MENU_NEW_TAB2
    
	Call ggoOper.SetReqAttr(frm1.txtDeptCd,   "N")
	Call ggoOper.SetReqAttr(frm1.txtDocCur,   "N")	
	Call ggoOper.SetReqAttr(frm1.txtTempGlDt, "N")
	'Call ggoOper.SetReqAttr(frm1.txtdesc,   "D")
    
    frm1.vspdData.MaxRows  = 0
    frm1.vspdData2.MaxRows = 0
    frm1.vspdData3.MaxRows = 0 
    
	SetGridFocus()
    SetGridFocus2()
    
	Call SetDefaultVal    
	Call InitVariables                                                      '⊙: Initializes local global variables	
	
	Call txtDocCur_OnChange()
		
    lgBlnFlgChgValue = False

    FncNew = True                              '⊙: Processing is OK
    lgFormLoad = True							' tempgldt read
    lgQueryOk = False
    lgstartfnc = False
End Function


'========================================================================================
Function FncDelete() 
	Dim IntRetCD 
    
    FncDelete = False           '⊙: Processing is NG
    Err.Clear                   '☜: Protect system from crashing
    lgBlnExecDelete = True
    'On Error Resume Next

	
    
    '-----------------------
    'Precheck area
    '-----------------------
    ' Update 상태인지를 확인한다.
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = False Then									'변경된 부분이 없을경우 

		intRetCd = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")				'삭제하시겠습니까?
		If intRetCd = VBNO Then
			Exit Function
		End IF

    Else

		IntRetCD = DisplayMsgBox("900038", Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
    	If IntRetCD = vbNo Then    		
      		Exit Function
    	End If
    End If

    '-----------------------
    'Delete function call area
    '-----------------------
    IF  DbDelete = False Then														'☜: Delete db data
    	Exit Function
    End If
    FncDelete = True 
    
End Function


'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                   '☜: Protect system from crashing
    
	'-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False  AND ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          'No data changed!!
        Exit Function
    End If    

    If CheckSpread3 = False then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")							'필수입력 check!!
        Exit Function
    End If
    
    If CheckSpread6 = False then
		IntRetCD = DisplayMsgBox("110420", "X", "X", "X")							'필수입력 check!!
        Exit Function
    End If
    
	If frm1.vspdData.MaxRows < 1 Then												'회계전표존재하지 않음 
		IntRetCD = DisplayMsgBox("114100", "X", "X", "X")
		Exit Function
	End If
  '-----------------------
    'Check content area
    '----------------------- 
    
    If Not chkField(Document, "2") Then                                  '⊙: Check contents area
		Exit Function
    End If
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If
    ggoSpread.Source = frm1.vspdData4
    If Not ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
		Exit Function
    End If
	'-----------------------
    'Save function call area
    '----------------------- 
    IF  DbSave	= False Then			                                                '☜: Save db data
		Exit Function
    End If
    
    FncSave = True                                                          
    
End Function


'========================================================================================
Function FncCopy() 

	Dim  IntRetCD
	 
	frm1.vspdData4.ReDraw = False	
	If frm1.vspdData4.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData4	
    ggoSpread.CopyRow
    
    frm1.vspdData4.Col = C_ItemSeq_2
	frm1.vspdData4.Row = frm1.vspdData4.ActiveRow
	lgIntMaxItemSeq = lgIntMaxItemSeq + 1
	frm1.vspdData4.value = lgIntMaxItemSeq
	
    Call SetSpreadColor2("I",0, frm1.vspdData4.ActiveRow, frm1.vspdData4.ActiveRow)
     Call SetSumItem()
    
	frm1.vspdData4.ReDraw = True
	Call vspdData4_Change(C_AcctCd, frm1.vspddata4.activerow)    
	
End Function


'========================================================================================================
Function FncCancel() 

	Dim iItemSeq
	
	If lgCurrentTabFg = C_TAB1 Then
		If frm1.vspdData.MaxRows < 1 Then Exit Function	

		With frm1.vspdData
		    .Row = .ActiveRow
		    .Col = 0		    
		    If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				IF len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					Call DeleteHSheet(.Text)
				end if	
		    End if       
			
		    ggoSpread.Source = frm1.vspdData	
		    ggoSpread.EditUndo
			Call SetSumItem()
			If .MaxRows = 0 Then
				Call SetToolbar(C_MENU_NEW_TAB1)
				lgPreToolBarTab1 = C_MENU_NEW_TAB1
				Exit Function
			End If

		    .Row = .ActiveRow
		    .Col = 0
		    
			if .Row = 0 then 			
				Exit Function
			end if
			
		    If .Text = ggoSpread.InsertFlag Then            
			    .Col = C_AcctCd
		        If Len(.Text) > 0 Then
					.Col = C_ItemSeq
					frm1.hItemSeq.value = .Text
		            frm1.vspdData2.MaxRows = 0
			        Call DbQuery3(.ActiveRow)
		        End If
		    Else
				.Col = C_ItemSeq
		        frm1.hItemSeq.value = .Text
		        frm1.vspdData2.MaxRows = 0
			    Call DbQuery2(.ActiveRow, C_VSPDDATA1)
		    End if		    
		End With
			
	ElseIf lgCurrentTabFg = C_TAB2 Then
		If frm1.vspdData4.MaxRows < 1 Then Exit Function	

		With frm1.vspdData4
		    .Row = .ActiveRow
		    .Col = 0		    
		    If .Text = ggoSpread.InsertFlag Then
				.Col = C_AcctCd
				IF len(Trim(.text)) > 0 Then 
					.Col = C_ItemSeq
					Call DeleteHSheet2(.Text)
				end if	
		    End if       
			
		    ggoSpread.Source = frm1.vspdData4	
		    ggoSpread.EditUndo
			Call SetSumItem()
			If .MaxRows = 0 Then
				Call SetToolbar(C_MENU_NEW_TAB2)
				lgPreToolBarTab1 = C_MENU_NEW_TAB2
				Exit Function
			End If

		    .Row = .ActiveRow
		    .Col = 0
		    
			if .Row = 0 then 			
				Exit Function
			end if
			
		    If .Text = ggoSpread.InsertFlag Then            
			    .Col = C_AcctCd
		        If Len(.Text) > 0 Then
					.Col = C_ItemSeq
					frm1.hItemSeq.value = .Text
		            frm1.vspdData5.MaxRows = 0
			        Call DbQuery4(.ActiveRow)
		        End If
		    Else
				.Col = C_ItemSeq
		        frm1.hItemSeq.value = .Text
		        frm1.vspdData5.MaxRows = 0
			    Call DbQuery2(.ActiveRow, C_VSPDDATA2)
		    End if		    
		End With		
    End If
End Function


'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
	Dim iRow
    Dim imRow
    Dim imRow2
    Dim exchRate
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
   '-----------------------
    'Check content area
    '----------------------- 
     If Not chkField(Document, "2") Then 
		Call ClickTab1()
        Exit Function
    End If

    If IsNumeric(Trim(pvRowCnt)) then
        imRow = CInt(pvRowCnt)
    Else

        imRow = AskSpdSheetAddRowCount()
       If imRow = "" Then
            Exit Function
        End If
    End If

    
	With frm1
	
		.vspdData4.focus
		ggoSpread.Source = .vspdData4	
		.vspdData4.ReDraw = False
		ggoSpread.InsertRow ,imRow
		
		Call SetSpreadColor2("I",0, frm1.vspdData4.ActiveRow, frm1.vspdData4.ActiveRow + imRow - 1)
		For iRow = .vspdData4.ActiveRow to .vspdData4.ActiveRow + imRow - 1
			'귀속부서를 default로 뿌려준다.
    	    .vspdData4.Row = iRow
			.vspdData4.col		= C_deptcd_2
			.vspdData4.value	= UCase  (.txtDeptCd.value)
			
			.vspdData4.col		= C_deptnm_2
			.vspdData4.value	= .txtDeptNm.value		
			
			.vspdData4.col		= C_ItemDesc_2
			.vspdData4.value	= .txtDesc.value
			
			.vspdData4.col		= C_DocCur_2
			.vspdData4.value	= parent.gcurrency
			
			.vspdData4.col		= C_ExchRate_2
			.vspdData4.value	= "1"		
					
			
			'입금전표이면 (01) '	'cr'을 넣어준다.
			IF  frm1.cboGlType.value = "01" Then
				.vspdData4.col = C_DrCrNm_2
				.vspdData4.value	= 1					
				.vspdData4.col = C_DrCrFg_2
				.vspdData4.value	= 1					
			ELSEIF frm1.cboGlType.value = "02" Then		
				.vspdData4.col = C_DrCrNm_2
				.vspdData4.value	= 2				
				.vspdData4.col = C_DrCrFg_2
				.vspdData4.value	= 2			
			END IF
			.vspdData4.Col = C_ItemSeq_2
			lgIntMaxItemSeq = lgIntMaxItemSeq + 1
			.vspdData4.value = lgIntMaxItemSeq
			.vspdData4.ReDraw = True
        Next
		
		
		.vspdData5.MaxRows =  0		        
        Call SetToolBar(C_MENU_CRT_TAB2)
        lgPreToolBarTab2 = C_MENU_CRT_TAB2
    End With
    Set gActiveElement = document.ActiveElement 
End Function


'========================================================================================
Function FncDeleteRow() 
	Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq

	If lgCurrentTabFg = C_TAB1 Then
		With frm1.vspdData 
			ggoSpread.Source = frm1.vspdData 
			.Row = .ActiveRow
			.Col = 0 		
			If frm1.vspdData.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function
			    .Col = 1 
			    DelItemSeq = .Text    	
				lDelRows = ggoSpread.DeleteRow    
		End With
		Call DeleteHsheet(DelItemSeq)
		    
    ElseIf lgCurrentTabFg = C_TAB2 Then
		With frm1.vspdData4 
			ggoSpread.Source = frm1.vspdData4 
			.Row = .ActiveRow
			.Col = 0 		
			If frm1.vspdData4.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function
			    .Col = 1 
			    DelItemSeq = .Text    	
				lDelRows = ggoSpread.DeleteRow    
		End With
		Call DeleteHsheet2(DelItemSeq)
	End If
    
    Call SetSumItem()
    
End Function


'========================================================================================
Function FncPrint() 
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function


'========================================================================================
Function FncExcel() 
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    Call parent.FncExport(Parent.C_MULTI)												'☜: 화면 유형 
End Function


'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase  (Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
Function FncExit()
	Dim IntRetCD
	Dim var1,var2
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then  
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")			'⊙: "Will you destory previous data"	
		If IntRetCD = vbNo Then
			Exit Function
		End If		
    end if    

    FncExit = True
    
End Function

'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	
	Call FncEBRPreview(ObjName,StrUrl)
	
End Function


'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId
    Dim StrEbrFile
    Dim intRetCd
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)
	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
	
   
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|TempGlNoFr|" & VarTempGlNoFr
	StrUrl = StrUrl & "|TempGlNoTo|" & VarTempGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function


'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, VarTempGlNoFr, VarTempGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a5118ma1"
	VarDateFr = UniConvDateToYYYYMMDD(frm1.txttempGlDt.Text, Parent.gDateFormat, Parent.gServerDateType)	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txttempGlDt.Text, Parent.gDateFormat, Parent.gServerDateType)

' 회계전표의 key는 GL_NO이기 때문에 GL_NO만 넘긴다.	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	VarTempGlNoFr = Trim(frm1.txttempGlNo.value)
	VarTempGlNoTo = Trim(frm1.txttempGlNo.value)
	varOrgChangeId = Trim(frm1.hOrgChangeId.value)
	
End Sub

'========================================================================================
Function FncBtnCalc() 

	Dim ii
	Dim tempAmt, tempLocAmt, tempExch, TempSep, tempDoc
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strDate
	Dim strExchFg
	Dim IntRetCD
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6		

	With frm1

		strSelect	= "b.minor_cd"
		strFrom		= "b_company a, b_minor b"
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If	

		strDate = UniConvDateToYYYYMMDD(frm1.txttempGLDt.text,parent.gDateFormat,"")

		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DocCur			
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_ItemAmt
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_ExchRate
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt		=	tempAmt * TempExch
					Else
						tempLocAmt		=	tempAmt / TempExch
					End If
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	tempLocAmt

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	tempAmt
				End If
			Next		
		End If

		If .vspdData4.MaxRows <> 0 Then
			For ii = 1 To .vspdData4.MaxRows
				.vspdData4.Row	=	ii
				.vspdData4.Col	=	C_DocCur_2			
				tempDoc			=	UCase(Trim(.vspdData4.text))
				.vspdData4.Col	=	C_ItemAmt_2
				tempAmt			=	UNICDbl(.vspdData4.text)
				.vspdData4.Col	=	C_ExchRate_2
				tempExch		=	UNICDbl(.vspdData4.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt		=	tempAmt * TempExch
					Else
						tempLocAmt		=	tempAmt / TempExch
					End If
					.vspdData4.Col	=	C_ItemLocAmt_2
					.vspdData4.text	=	tempLocAmt

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData4.Col	=	C_ItemLocAmt_2
					.vspdData4.text	=	tempAmt
				End If
			Next		
		End If
	End With
	Call SetSumItem	
End Function


'==========================================================================================
Sub DocCur_OnChange(byVal strDocCur, byVal Row)

    lgBlnFlgChgValue = True
	If Trim(strDocCur) = parent.gCurrency Then
		frm1.vspdData4.Col  = C_ExchRate_2
		frm1.vspdData4.Text = "1"
	Else
		Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txttempGLDt.text,parent.gDateFormat,""), UCase(Trim(strDocCur)),Row)
	End IF
	Call SetSumItem()

End Sub

'=======================================================================================================
'   Function Name : FindExchRate
'   Function Desc : 1.날짜, Row를 입력받아 날짜에 해당하는 환율정보를 읽어온다.
'=======================================================================================================

Function FindExchRate(Byval strDate, Byval FromCurrency,Byval Row )
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strExchFg
	Dim strExchRate
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
	Dim IntRetCD		

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
			frm1.vspdData4.Col  = C_ExchRate_2
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
			frm1.vspdData4.Col  = C_ExchRate_2
			frm1.vspdData4.Text = UNIConvNumPCToCompanyByCurrency(strExchRate, parent.gCurrency, parent.ggExchRateNo, parent.gLocRndPolicyNo, "X")
		Else
			IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
		End If
	End If
	
End Function    


'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim RetFlag

    DbQuery = False
    Call LayerShowHide(1)
    frm1.vspdData3.MaxRows = 0 

    Err.Clear                '☜: Protect system from crashing
    
    With frm1    				
		
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtTempGlNo=" & UCase  (Trim(.htxtTempGlNo.value))	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 
			strVal = strVal & "&txtTempGlNo=" & UCase  (Trim(.txtTempGlNo.value))	'조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag
		End If
   
		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인   
   
		Call RunMyBizASP(MyBizASP, strVal)		'☜: 비지니스 ASP 를 가동 

    End With
    
    DbQuery = True

End Function

'=======================================================================================================
Function DbQueryOk()

	Dim iIntRow
	Dim iIntIndex
	
	Call ClickTab1()
	
	With frm1
		   
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = Parent.OPMD_UMODE												'Indicates that current mode is Update mode        
		.txtCommandMode.value = "UPDATE"
				
		For iIntRow = 1 To .vspdData.MaxRows					
			.vspdData.Row = iIntRow
			.vspdData.Col = C_DrCrFg
			iIntIndex = .vspdData.value
			.vspdData.col = C_DrCrNm
			.vspdData.value = iIntIndex					
		Next
		
		For iIntRow = 1 To .vspdData4.MaxRows					
			.vspdData4.Row = iIntRow
			.vspdData4.Col = C_DrCrFg_2
			iIntIndex = .vspdData4.value
			.vspdData4.col = C_DrCrNm_2
			.vspdData4.value = iIntIndex					
		Next		
					
		Call ggoOper.SetReqAttr(frm1.txtTempGlDt,	"Q")
		Call ggoOper.SetReqAttr(frm1.cboGlType,		"Q")
		Call ggoOper.SetReqAttr(frm1.txtDocCur,		"Q")
		
		If frm1.cboGlInputType.Value <> C_GLINPUTTYPE then
			Call SetToolbar(C_MENU_NEW_TAB1)
			lgPreToolBarTab1 = C_MENU_NEW_TAB1
			lgPreToolBarTab2 = C_MENU_NEW_TAB1
		Else
			Call SetToolbar(C_MENU_UPD_TAB1)
			lgPreToolBarTab1 = C_MENU_UPD_TAB1
			lgPreToolBarTab2 = C_MENU_UPD_TAB2
		End If
			
		Call SetSpreadLock("Q", 1, 1, "")			
		Call SetSpread2Lock("Q", 1, 1, "")
		Call SetSpread4Lock("Q", 1, 1, "")			
		Call SetSpread5Lock("Q", 1, 1, "")
		Call ggoOper.SetReqAttr(frm1.txtDeptCd,	"Q")													 
		Call ggoOper.SetReqAttr(frm1.txtdesc,   "Q")	
	
		If .vspdData.MaxRows > 0 Then
			.vspdData.Row = 1
			.vspdData.Col = 1
			.hItemSeq.Value = .vspdData.Text
			Call DbQuery2(1, C_VSPDDATA1)
		End If
		
		If .vspdData4.MaxRows > 0 Then
			.vspdData4.Row = 1
			.vspdData4.Col = 1
			Call DbQuery2(1, C_VSPDDATA4)
		End If
    
    End With
    
    call txtDocCur_OnChange()
    Call txtDeptCd_OnChange()
    SetGridFocus()
    SetGridFocus2()		
    
    lgBlnFlgChgValue = False
    Call CancelRestoreToolBar()
    Set gActiveElement = document.ActiveElement 

End Function

							
'=======================================================================================================
Function DbQuery2(ByVal Row, ByVal pVspdData)

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
	Dim i
	Dim arrVal
	
	Dim iStrTempGlNo
	Dim iStrTempGlItemSeq
	Dim iStrOpenGlNo
	Dim iStrOpenGlItemSeq
	Dim iStrGlNo
	Dim iStrItemSeq
	Dim iObjSpread1
	Dim iObjSpread2
	Dim iObjSpread3
	
	If pVspdData = C_VSPDDATA1 Then
	
		Set iObjSpread1 = frm1.vspdData
		Set iObjSpread2 = frm1.vspdData2
		Set iObjSpread3 = frm1.vspdData3
	ElseIf pVspdData = C_VSPDDATA4 Then
		Set iObjSpread1 = frm1.vspdData4
		Set iObjSpread2 = frm1.vspdData5
		Set iObjSpread3 = frm1.vspdData6
	End If
		
	With frm1
	  If pVspdData = C_VSPDDATA1 Then	
		iStrTempGlNo		= frm1.txtTempGlNo.value	'나중에 .htxtTempGlNo.value로 바꾸장 
	    iObjSpread1.Row		= Row
	    iObjSpread1.Col		= C_ItemSeq
	    iStrTempGlItemSeq	= Trim(iObjSpread1.Text)
	    iObjSpread1.Col		= C_OpenGlNo
	    iStrOpenGlNo		= iObjSpread1.Text
	    iObjSpread1.Col		= C_OpenGlItemSeq
	    iStrOpenGlItemSeq	= iObjSpread1.Text
	  ElseIf pVspdData = C_VSPDDATA4 Then
	    iStrTempGlNo		= frm1.txtTempGlNo.value	'나중에 .htxtTempGlNo.value로 바꾸장 
		iObjSpread1.Row		= Row
		iObjSpread1.Col		= C_ItemSeq_2
		iStrTempGlItemSeq	= Trim(iObjSpread1.Text)
		iObjSpread1.Col		= C_OpenGlNo_2
		iStrOpenGlNo		= iObjSpread1.Text
		iObjSpread1.Col		= C_OpenGlItemSeq_2
		iStrOpenGlItemSeq	= iObjSpread1.Text  
	  End If  
	    	    
	    iObjSpread2.ReDraw = false	
	    If pVspdData = C_VSPDDATA1 Then
			If CopyFromData(iStrTempGlItemSeq) = True Then					
				If lgIntFlgMode = Parent.OPMD_CMODE Then					
					Call SetSpread2Lock("","1","1","")
				Else
					Call SetSpread2Lock("","1","1","")
				End If				
				Exit Function
			End If
		ElseIf pVspdData = C_VSPDDATA4 Then
			If CopyFromData2(iStrTempGlItemSeq) = True Then	
				If lgIntFlgMode = Parent.OPMD_CMODE Then					
					Call SetSpread4Color()
				Else
					Call SetSpread4Color()
					Call CtrlSpreadLock2("","",1,-1)
				End If				
				Exit Function
			End If
		End If
	         
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		iObjSpread1.Row = Row
		
		If pVspdData = C_VSPDDATA1 Then
		   iObjSpread1.Col = C_ItemSeq
		Else   
		   iObjSpread1.Col = C_ItemSeq_2
		End If   
		   
		
		If iStrOpenGlNo <> "" And iStrOpenGlItemSeq <> "" Then	
				
			strSelect =				" C.DTL_SEQ, "		
			strSelect = strSelect & " A.CTRL_CD, "
			strSelect = strSelect & " A.CTRL_NM, "
			strSelect = strSelect & " C.CTRL_VAL, "
			strSelect = strSelect & " '', "		
			strSelect = strSelect & " Case    WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END, "
			strSelect = strSelect &	  iStrOpenGlItemSeq  & ", "
			strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
			strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
			strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
			strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
			strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
			strSelect = strSelect & " Case  	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
			strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
			strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  END, " & iStrOpenGlItemSeq & ","
			strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
			strFrom	=			" A_CTRL_ITEM  		A (NOLOCK), "
			strFrom = strFrom & " A_ACCT_CTRL_ASSN  	B (NOLOCK), "
			strFrom = strFrom & " A_GL_DTL  			C (NOLOCK), "
			strFrom = strFrom & " A_GL_ITEM  			D (NOLOCK)	"
					
			strWhere =			  " D.GL_NO = " & FilterVar(UCase  (Trim(iStrOpenGlNo)), "''", "S")   
			strWhere = strWhere & " AND D.ITEM_SEQ	= " & iStrOpenGlItemSeq & " "
			strWhere = strWhere & " AND D.GL_NO		=  C.GL_NO  "
			strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
			strWhere = strWhere & "	AND D.ACCT_CD	*= B.ACCT_CD "
			strWhere = strWhere & " AND C.CTRL_CD	*= B.CTRL_CD "		
			strWhere = strWhere & " AND C.CTRL_CD	= A.CTRL_CD "
			strWhere = strWhere & " ORDER BY C.DTL_SEQ "	
		
		Else			
			strSelect =				" C.DTL_SEQ, "		
			strSelect = strSelect & " A.CTRL_CD, "
			strSelect = strSelect & " A.CTRL_NM, "
			strSelect = strSelect & " C.CTRL_VAL, "
			strSelect = strSelect & " '', "		
			strSelect = strSelect & " Case    WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END, "
			strSelect = strSelect &	  iStrTempGlItemSeq  & ", "
			strSelect = strSelect & " LTrim(ISNULL(A.TBL_ID,'')), "		
			strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_ID,'')), "
			strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')), "		
			strSelect = strSelect & " LTrim(ISNULL(A.COLM_DATA_TYPE,'')), "
			strSelect = strSelect & " LTrim(ISNULL(A.DATA_LEN,'')), "
			strSelect = strSelect & " Case  	WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
			strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
			strSelect = strSelect & "		WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  END, " & iStrTempGlItemSeq & ","	
			strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    	
			strFrom	=			" A_CTRL_ITEM  		A (NOLOCK), "
			strFrom = strFrom & " A_ACCT_CTRL_ASSN  	B (NOLOCK), "
			strFrom = strFrom & " A_GL_DTL  		C (NOLOCK), "
			strFrom = strFrom & " A_GL_ITEM  	D (NOLOCK)	"
					
			strWhere =			  " D.GL_NO = " & FilterVar(UCase  (Trim(iStrTempGlNo)), "''", "S")   
			strWhere = strWhere & " AND D.ITEM_SEQ		= " & iStrTempGlItemSeq & " "
			strWhere = strWhere & " AND D.GL_NO			=  C.GL_NO  "
			strWhere = strWhere & " AND D.ITEM_SEQ		=  C.ITEM_SEQ "
			strWhere = strWhere & "	AND D.ACCT_CD		*= B.ACCT_CD "
			strWhere = strWhere & " AND C.CTRL_CD		*= B.CTRL_CD "		
			strWhere = strWhere & " AND C.CTRL_CD		= A.CTRL_CD "
			strWhere = strWhere & " ORDER BY C.DTL_SEQ "
		
		End If
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = iObjSpread2
			ggoSpread.SSShowData lgF2By2

			Select Case   pVspdData
				Case   C_VSPDDATA1
			
					For lngRows = 1 To iObjSpread2.Maxrows
						iObjSpread2.row = lngRows	
						iObjSpread2.col = C_Tableid 
						IF Trim(iObjSpread2.text) <> "" Then
							iObjSpread2.col = C_Tableid
							strTableid = iObjSpread2.text
							iObjSpread2.col = C_Colid
							strColid = iObjSpread2.text
							iObjSpread2.col = C_ColNm
							strColNm = iObjSpread2.text	
							iObjSpread2.col = C_MajorCd					
							strMajorCd = iObjSpread2.text	
							
							iObjSpread2.col = C_CtrlVal
							
							strNmwhere = strColid & " =  " & FilterVar(UCase  (Trim(iObjSpread2.text)), "''", "S")
							
							IF Trim(strMajorCd) <> "" Then
								strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
							End IF				 
							
							IF CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
								iObjSpread2.col = C_CtrlValNm
								arrVal = Split(lgF0, Chr(11))  
								iObjSpread2.text = arrVal(0)
							End IF
						End IF								
						
						strVal = strVal & Chr(11) & iStrTempGlItemSeq
						
						iObjSpread2.Col = C_DtlSeq
						strVal = strVal & Chr(11) & iObjSpread2.Text
						
						iObjSpread2.Col = C_CtrlCd
						strVal = strVal & Chr(11) & iObjSpread2.Text
								
						iObjSpread2.Col = C_CtrlNm
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_CtrlVal
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_CtrlPB
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_CtrlValNm
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_Seq
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_Tableid
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_Colid
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_ColNm
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_Datatype
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_DataLen
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_DRFg
						strVal = strVal & Chr(11) & iObjSpread2.Text
								    
						iObjSpread2.Col = C_MajorCd
						strVal = strVal & Chr(11) & iObjSpread2.Text

						.vspdData2.Col = C_MajorCd + 1
		'				.vspdData2.Text = lngRows
						strVal = strVal & Chr(11) & lngRows

						strVal = strVal & Chr(11) & Chr(12)			
					NEXT					
				Case   C_VSPDDATA4
					For lngRows = 1 To iObjSpread2.Maxrows
						iObjSpread2.row = lngRows	
						iObjSpread2.col = C_Tableid_2 
						IF Trim(iObjSpread2.text) <> "" Then
							iObjSpread2.col = C_Tableid_2
							strTableid = iObjSpread2.text
							iObjSpread2.col = C_Colid_2
							strColid = iObjSpread2.text
							iObjSpread2.col = C_ColNm_2
							strColNm = iObjSpread2.text	
							iObjSpread2.col = C_MajorCd_2					
							strMajorCd = iObjSpread2.text	
								
							iObjSpread2.col = C_CtrlVal
								
							strNmwhere = strColid & " =  " & FilterVar(UCase  (Trim(iObjSpread2.text)), "''", "S")
								
							IF Trim(strMajorCd) <> "" Then
								strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
							End IF				 
								
							IF CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
								iObjSpread2.col = C_CtrlValNm
								arrVal = Split(lgF0, Chr(11))  
								iObjSpread2.text = arrVal(0)
							End IF
						End IF												
					
						strVal = strVal & Chr(11) & iStrTempGlItemSeq

						iObjSpread2.Col = C_DtlSeq_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							
						iObjSpread2.Col = C_CtrlCd_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							
						iObjSpread2.Col = C_CtrlNm_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_CtrlVal_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_CtrlPB_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_CtrlValNm_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Seq_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Tableid_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Colid_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_ColNm_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_Datatype_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_DataLen_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_DRFg_2
						strVal = strVal & Chr(11) & iObjSpread2.Text
							    
						iObjSpread2.Col = C_MajorCd_2
						strVal = strVal & Chr(11) & iObjSpread2.Text

						.vspdData2.Col = C_MajorCd_2 + 1
		'				.vspdData2.Text = lngRows
						strVal = strVal & Chr(11) & lngRows

						strVal = strVal & Chr(11) & Chr(12)								
					Next						
				Case   Else
			End Select		
		

			ggoSpread.Source = iObjSpread3			
			ggoSpread.SSShowData strVal	
			
		END IF 		
		

		intItemCnt = iObjSpread1.MaxRows
        
        If pVspdData = C_VSPDDATA1 Then				
			If lgIntFlgMode = Parent.OPMD_CMODE Then					
				Call SetSpread2Lock("","1","1","")
			Else
				Call SetSpread2Lock("","1","1","")
			End If
		ElseIf pVspdData = C_VSPDDATA4 Then		
			If lgIntFlgMode = Parent.OPMD_CMODE Then					
				Call SetSpread4Color()
			Else					
				Call SetSpread4Color()									
				Call CtrlSpreadLock2("","",1,1)
			End If				
		End If
		
		
	End With
	
	iObjSpread2.ReDraw = True
	
	Call LayerShowHide(0)
	
	DbQuery2 = True
	lgQueryOk = True
End Function

'========================================================================================================
Sub InitData(ByVal pVspdData)
	Dim intRow
	Dim intIndex 
	If pVspdData = C_VSPDDATA1 Then
		With frm1.vspdData
			For intRow = 1 To .MaxRows
				.Row = intRow
				.col = C_DrCrFg     :			intIndex = .value
				.col = C_DrCrNm     :			.value = intindex
						
			Next	
		End With
	ElseIf pVspdData = C_VSPDDATA4 Then
		With frm1.vspdData4
			For intRow = 1 To .MaxRows
				
				.Row = intRow

				.col = C_DrCrFg_2     :			intIndex = .value
				.col = C_DrCrNm_2     :			.value = intindex
						
			Next	
		End With
	End If

End Sub

'========================================================================================================
Function DbSave() 
    
    Dim pAP010M 
    Dim lngRows , itemRows
    Dim lGrpcnt
    DIM strVal 
    Dim tempItemSeq
	Dim	intRetCd
	Dim ii	
    Dim strNote
    Dim strItemDesc
    Dim iColSep
    Dim iRowSep     
    
    strNote = ""
    DbSave = False
    
    Call LayerShowHide(1)
    
	With frm1
		.txtFlgMode.value     = lgIntFlgMode
		.txtUpdtUserId.value  = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtMode.value        = Parent.UID_M0002
		.txtAuthorityFlag.value     = lgAuthorityFlag               '권한관리 추가	
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
   
    lGrpCnt = 1
    strVal = ""
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	
   
    
    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
	    
		For lngRows = 1 To .MaxRows
    
			.Row = lngRows
			.Col = 0

			If .Text <> ggoSpread.DeleteFlag Then
			
				'CRUD
				strVal = strVal & "C" & iColSep
				
				'CurrentRow 
				strVal = strVal & lngRows & iColSep
								
			    .Col = C_ItemSeq	
			    strVal = strVal & Trim(.Text) & iColSep

				.Col = C_AcctCd		
			    strVal = strVal & Trim(.Text) & iColSep
				
				.Col = C_DrCrFG		
			    strVal = strVal & Trim(.Text) & iColSep
			    
			    'OrgChangeId
			   .Col = C_OrgChangeId
			    strVal = strVal & Trim(.Text) & iColSep
				
				.Col = C_deptcd	    
			    strVal = strVal & Trim(.Text) & iColSep
			    
			    'DocCur
			    strVal = strVal & UCase(frm1.txtDocCur.value) & iColSep
			    
			    .Col = C_ExchRate	
			    strVal = strVal & UNICDbl(Trim(.Text)) & iColSep
			    
			    'VarType
				strVal = strVal & "" & iColSep
			        
			    .Col = C_ItemAmt	
			    strVal = strVal & UNICDbl(Trim(.Text)) & iColSep
				
				'.Col = C_IsLAmtChange	
				
				'Local 금액을 사용자 입력시 입력금액을 전달 
				'If .Text = "Y" Then
   					.Col = C_ItemLocAmt	'6
					strVal = strVal & UNICDbl(Trim(.Text)) & iColSep
				'Else
				'	strVal = strVal & "0" & iColSep
				'End If

			    .Col = C_ItemDesc	'7
			    strItemDesc = Trim(.Text)
			    
			    If Trim(strItemDesc) = "" Or isnull(strItemDesc) Then
					 ggoSpread.Source = frm1.vspdData3
					'----------------------------------------------------
					 frm1.vspdData.Col = 1
					 tempItemSeq = frm1.vspdData.Text  
					 strNote = ""
					 With frm1.vspdData3
							For itemRows = 1 to frm1.vspdData3.MaxRows
								.Row = itemRows
								.Col = 1
								
								if .Text =  tempItemSeq then 					
									.Col= 9 'C_Tableid	+ 1				
									IF 	.Text = "B_BIZ_PARTNER" OR .Text = "B_BANK" OR .Text = "F_DPST" THEN
										.Col = 7 'C_CtrlValNm + 1 
									ELSE
										.Col = 5 'C_CtrlVal + 1 
									END IF	
									strNote = strNote & C_NoteSep & Trim(.Text)
								end if		    
							Next
							strNote = Mid(strNote,2)
					 End With
					 
					 strVal = strVal & strNote & iColSep
						'----------------------------------------------------
					 ggoSpread.Source = frm1.vspdData
			    Else
					strVal = strVal & strItemDesc & iColSep
			    End if
			    
   			    .Col = C_GL_DT
			    strVal = strVal & UniConvDate(Trim(.Text)) & iColSep
			    
			    .Col = C_OpenGlNo
				strVal = strVal & Trim(.Text) & iColSep
			    
			    .Col = C_OpenGlItemSeq
			    strVal = strVal & Trim(.Text) & iColSep
			    			    
			    .Col = C_MgntFg
			    strVal = strVal & Trim(.Text) & iRowSep	  
			      	
			End If		
		Next

    End With
	
    frm1.txtSpread.value  = strVal									'Spread Sheet 내용을 저장    
	
	IF frm1.txtSpread.value = "" Then	
		intRetCd = DisplayMsgBox("990008", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		If intRetCd = VBNO Then
			Exit Function
		End IF	
		Call DbDelete
		ggoSpread.Source = frm1.vspdData
        ggoSpread.ClearSpreadData
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
        ggoSpread.Source = frm1.vspdData3
        ggoSpread.ClearSpreadData
		Call InitVariables	    
		Exit Function
	END IF
			
    
    strVal = ""    
    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl 저장 
    
		For itemRows = 1 To frm1.vspdData.MaxRows 
 		    frm1.vspdData.Row = itemRows
		    frm1.vspdData.Col = 0

		    if frm1.vspdData.Text <> ggoSpread.DeleteFlag then	

				frm1.vspdData.Col = C_ItemSeq
			    tempItemSEq = frm1.vspdData.Text  
		        
			    For lngRows = 1 To .MaxRows
			    
					.Row = lngRows
					.Col = C_ItemSeq

					IF .text = tempitemseq THEN
						.Col = 0 					
						strVal = strVal & "C" & iColSep
						
						.Col = 1 		 			'ItemSEQ						        
						strVal = strVal & tempitemseq & iColSep
						
						.Col =  2 'C_DtlSeq + 1   				'Dtl SEQ					        
						strVal = strVal & Trim(.Text) & iColSep
						
						.Col =  3 'C_CtrlCd + 1		 		'관리항목코드							
						strVal = strVal & Trim(.Text) & iColSep
						
						.Col = 5 'C_CtrlVal + 1				'관리항목 Value 					        
						strVal = strVal & UCase(Trim(.Text)) & iRowSep	
						
					End IF			
		    	Next
		   End If
   		Next

    End With
    
    frm1.txtSpread3.value  = strVal						'Spread Sheet 내용을 저장 
    
    
    
    '======================================================================================
    
    
    strVal = ""    
    With frm1.vspdData4
	    
		For lngRows = 1 To .MaxRows
    
			.Row = lngRows
			.Col = 0

			If .Text <> ggoSpread.DeleteFlag Then
				'CRUD
				strVal = strVal & "C" & iColSep
				
				'CurrentRow 
				strVal = strVal & lngRows & iColSep
								
			    .Col = C_ItemSeq_2	
			    strVal = strVal & Trim(.Text) & iColSep

				.Col = C_AcctCd_2	
			    strVal = strVal & Trim(.Text) & iColSep
				
				.Col = C_DrCrFG_2		
			    strVal = strVal & Trim(.Text) & iColSep
			    
			    'OrgChangeId
			    strVal = strVal & frm1.hOrgChangeId.value & iColSep
				
				.Col = C_deptcd_2	    
			    strVal = strVal & Trim(.Text) & iColSep
			    
			    'DocCur
   			    .Col = C_DocCur_2
			    strVal = strVal & UCase(Trim(.Text)) & iColSep
			    
			    .Col = C_ExchRate_2	
			    strVal = strVal & UNICDbl(Trim(.Text)) & iColSep
			    
			    'VarType
				strVal = strVal & "" & iColSep
				
			    .Col = C_ItemAmt_2	
			    strVal = strVal & UNICDbl(Trim(.Text)) & iColSep
				
				
				'Local 금액을 사용자 입력시 입력금액을 전달 

   					.Col = C_ItemLocAmt_2	'6
					strVal = strVal & UNICDbl(Trim(.Text)) & iColSep

			    .Col = C_ItemDesc_2	'7
			    strItemDesc = Trim(.Text)
			    
			    If Trim(strItemDesc) = "" Or isnull(strItemDesc) Then
					 ggoSpread.Source = frm1.vspdData3
					'----------------------------------------------------
					 frm1.vspdData4.Col = C_ItemSeq_2
					 tempItemSeq = frm1.vspdData4.Text  
					 strNote = ""
					 With frm1.vspdData6
							For itemRows = 1 to frm1.vspdData6.MaxRows
								.Row = itemRows
								.Col = C_ItemSeq_2
								
								if .Text =  tempItemSeq then 					
									.Col= 9 'C_Tableid	+ 1				
									IF 	.Text = "B_BIZ_PARTNER" OR .Text = "B_BANK" OR .Text = "F_DPST" THEN
										.Col = 7 'C_CtrlValNm + 1 
									ELSE
										.Col = 5 'C_CtrlVal + 1 
									END IF	
									strNote = strNote & C_NoteSep & Trim(.Text)
								end if		    
							Next
							strNote = Mid(strNote,2)
					 End With
					 
					 strVal = strVal & strNote & iColSep
						'----------------------------------------------------
					 ggoSpread.Source = frm1.vspdData4
			    Else
					strVal = strVal & strItemDesc & iColSep
			    End if
			    
   				strVal = strVal & "" & iColSep	    
    
			   
				.Col = C_OpenGlNo_2
				strVal = strVal & Trim(.Text) & iColSep
			    
			     .Col = C_OpenGlItemSeq_2
			    strVal = strVal & Trim(.Text) & iColSep
			    			    
			    .Col = C_MgntFg_2
			    strVal = strVal & Trim(.Text) & iRowSep	 
			   
			End If		
		Next

    End With
	
   
    frm1.txtSpread.value  = frm1.txtSpread.value & strVal									'Spread Sheet 내용을 저장    

	IF frm1.txtSpread.value = "" Then	
		intRetCd = DisplayMsgBox("990008", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		If intRetCd = VBNO Then
			Exit Function
		End IF	
		Call DbDelete
		ggoSpread.Source = frm1.vspdData4
        ggoSpread.ClearSpreadData
        ggoSpread.Source = frm1.vspdData5
        ggoSpread.ClearSpreadData
        ggoSpread.Source = frm1.vspdData6
        ggoSpread.ClearSpreadData

		Call InitVariables	    
		Exit Function
	END IF
			
    
    strVal = ""
    ggoSpread.Source = frm1.vspdData6

    With frm1.vspdData6      ' Dtl 저장 
    
		For itemRows = 1 To frm1.vspdData4.MaxRows 
 		    frm1.vspdData4.Row = itemRows
		    frm1.vspdData4.Col = 0

		    if frm1.vspdData4.Text <> ggoSpread.DeleteFlag then	

				frm1.vspdData4.Col = C_ItemSeq_2
			    tempItemSEq = frm1.vspdData4.Text  
		        
			    For lngRows = 1 To .MaxRows
			    
					.Row = lngRows
					.Col = C_ItemSeq_2
					
					IF .text = tempitemseq THEN
					
						.Col = 0 						
						strVal = strVal & "C" & iColSep
						
						.Col = 1 		 			'ItemSEQ							        
						strVal = strVal & tempitemseq & iColSep
						
						.Col = 2 'C_DtlSeq + 1   				'Dtl SEQ						        
						strVal = strVal & Trim(.Text) & iColSep
						
						.Col = 3 'C_CtrlCd + 1		 		'관리항목코드								
						strVal = strVal & Trim(.Text) & iColSep
						
						.Col = 5 'C_CtrlVal + 1				'관리항목 Value 						        
						strVal = strVal & UCase  (Trim(.Text)) & iRowSep	
						
					End IF			
		    	Next
		   End If
   		Next

    End With

    With frm1
		.txtSpread3.value  = frm1.txtSpread3.value & strVal						'Spread Sheet 내용을 저장 
    		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
	End With	
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)											'저장 비지니스 ASP 를 가동 
        
    DbSave = True                                                           
    
End Function

'========================================================================================
Function DbSaveOk(ByVal TempGlNo)					'☆: 저장 성공후 실행 로직 

	lgBlnFlgChgValue = false
	
	frm1.txtTempGlNo.value = UCase  (Trim(TempGlNo))
    frm1.txtCommandMode.value = "UPDATE"
    
	Call ggoOper.ClearField(Document, "2")      '⊙: Condition field clear    
    Call InitVariables							'⊙: Initializes local global variables
	lgPreToolBarTab1 = C_MENU_UPD_TAB1
	Call DbQuery
	
End Function

'========================================================================================
Function DbDelete()
	Dim strVal
	
    Err.Clear
    Call LayerShowHide(1)    
	DbDelete = False														'⊙: Processing is NG

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003						'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal & "&txtTempGlNo=" & UCase(Trim(frm1.txtTempGlNo.value))	'☜: 삭제 조건 데이타 
    strVal = strVal & "&txtDeptCd=" & UCase(Trim(frm1.txtDeptCd.value))
	strVal = strVal & "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.value)
	strVal = strVal & "&txtTempGlDt=" & Trim(frm1.txttempgldt.text)

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True                                                         '⊙: Processing is NG	
    
End Function

'=======================================================================================================
Function DbDeleteOk()													'삭제 성공후 실행 로직 
	Call FncNew()	
End Function



'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()


End Sub


'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_ItemAmt,-1, .txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		
	End With

End Sub
    
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub

'=======================================================================================================  
Sub InputCtrlVal(ByVal pRow, ByVal pVspdData)
	Dim iObjSpread1
	Dim iObjSpread2
	Dim strAcctCd		
	Dim ii

	If pVspdData = C_VSPDDATA1 Then
		Set iObjSpread1 = frm1.vspdData
		Set iObjSpread2 = frm1.vspdData2
	ElseIf pVspdData = C_VSPDDATA4 Then
		Set iObjSpread1 = frm1.vspdData4
		Set iObjSpread2 = frm1.vspdData5
	Else
		Exit Sub
	End If
		
	lgBlnFlgChgValue = True

	If pVspdData = C_VSPDDATA1 Then
		ggoSpread.Source = iObjSpread1
		iObjSpread1.Col = C_AcctCd
		iObjSpread1.Row = pRow	
		strAcctCd	= Trim(iObjSpread1.text)		
		
		iObjSpread1.Col = C_deptcd
		iObjSpread1.Row = pRow		
	ElseIf pVspdData = C_VSPDDATA4 Then
		ggoSpread.Source = iObjSpread1
		iObjSpread1.Col = C_AcctCd_2
		iObjSpread1.Row = pRow	
		strAcctCd	= Trim(iObjSpread1.text)		
		
		iObjSpread1.Col = C_deptcd_2
		iObjSpread1.Row = pRow		
	End If
		
	If pVspdData = C_VSPDDATA1 Then
		Call AutoInputDetail(strAcctCd, Trim(iObjSpread1.text), frm1.txttempGLDt.text, pRow)			
		iObjSpread2.Col = C_CtrlVal
		For ii = 1 To iObjSpread2.MaxRows
			iObjSpread2.Row = ii					
			If Trim(iObjSpread2.text) <> "" Then
				Call CopyToHSheet2(iObjSpread1.ActiveRow,ii)			 			
			End if
		Next
	ElseIf pVspdData = C_VSPDDATA4 Then
		Call AutoInputDetail2(strAcctCd, Trim(iObjSpread1.text), frm1.txttempGLDt.text, pRow)
		iObjSpread2.Col = C_CtrlVal_2
		For ii = 1 To iObjSpread2.MaxRows
			iObjSpread2.Row = ii					
			If Trim(iObjSpread2.text) <> "" Then
				Call CopyToHSheet4(iObjSpread1.ActiveRow,ii)			 			
			End if
		Next
	End If
End Sub	

    
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

-->
</SCRIPT>
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>미결정리내역</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupOA()">미결연결</A>&nbsp;|&nbsp;
											<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=10>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH="100%"> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%"> 
					    <FIELDSET CLASS="CLSFLD">
						  <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>전표번호</TD>
								<TD CLASS=TD656 NOWRAP><INPUT NAME="txtTempGlNo" ALT="전표번호" MAXLENGTH="18" SIZE=20 STYLE="TEXT-ALIGN: left" tag  ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefTempGl()"></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%" SCROLL="no">
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>								
								<TD CLASS=TD5 NOWRAP>전표일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txttempGLDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="전표일자" tag="22" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>전표형태</TD>								
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlType" tag="24" STYLE="WIDTH:82px:" ALT="전표형태"><OPTION VALUE="" selected></OPTION></SELECT></TD> 
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value)" tag="22">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X"></TD>
													 <INPUT NAME="txtInternalCd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value, C_POPUP_DOCCUR, C_CONDFIELD)"></TD>
						   </TR>
						   <TR>									
								<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="24" STYLE="WIDTH:82px:" ALT="전표입력경로"><OPTION VALUE="" selected></OPTION></SELECT></TD>																
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>						
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="비고" MAXLENGTH="128" SIZE="70" tag="22N" ></TD>
							</TR>	
							<TR>
								<TD HEIGHT="60%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTTON>&nbsp;
								<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
								<TD CLASS=TD6 NOWRAP>
								&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(자국)" id=OBJECT3></OBJECT>');</SCRIPT>
								&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="대변합계(자국)" id=OBJECT4></OBJECT>');</SCRIPT>
								</TD>
							</TR>
			                <TR>
								<TD HEIGHT="40%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>
					</DIV>
					<!-- 두번째 탭 내용 -->  
					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="66%">
								<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData4 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
								
							</TR>
							<TR>
								<TD COLSPAN=4>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR>														
											<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTTON>&nbsp;
											<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
											<TD CLASS=TD6 NOWRAP>
											&nbsp;
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="차변합계(자국)" id=OBJECT3></OBJECT>');</SCRIPT>
											&nbsp;
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt2 style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="대변합계(자국)" id=OBJECT4></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>																					
							</TR>
						    <TR HEIGHT="34%">
								<TD WIDTH="100%" COLSPAN="4">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData5 width="100%" tag="2" TITLE="SPREAD" id=OBJECT6> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread3><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TABINDEX="-1" CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData6 WIDTH=0 HEIGHT=0 tag="23" TITLE="SPREAD" id=vaSpread6><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<TEXTAREA class=hidden name=txtSpread		tag="24" Tabindex="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" Tabindex="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="htxtTempGlNo"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtCommandMode"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" Tabindex="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"  tag="24" Tabindex="-1"><!--권한관리추가 -->
<INPUT TYPE=HIDDEN NAME="hCongFg"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>
</BODY>
</HTML>
