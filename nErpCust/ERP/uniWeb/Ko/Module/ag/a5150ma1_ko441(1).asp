<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : Multi Receipt
'*  3. Program ID           : a3130ma1.asp
'*  4. Program Name         : 통합반제 
'*  5. Program Desc         :
'*  6. Complus List         : PARG100
'*  7. Modified date(First) : 2006/09/12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☆) Means that "must change"
'* 13. HisTory              :
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
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../ag/Acctctrl_ko441_1.vbs">				</SCRIPT>

<SCRIPT LANGUAGE=vbscript>
Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const BIZ_PGM_ID  = "a5150mb1_ko441.asp"												'☆: 비지니스 로직 ASP명 

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'☆: 환율정보 비지니스 로직 ASP명 

Const TAB1 = 1																		'☜: Tab의 위치 
Const TAB2 = 2

Dim C_OPEN_TYPE
Dim C_OPEN_TYPE_NM
Dim C_GL_NO
Dim C_OPEN_DT
Dim C_BP_CD
Dim C_BP_NM
Dim C_OPEN_ACCT_CD
Dim C_OPEN_ACCT_NM
Dim C_DR_CR_FG
Dim C_DR_CR_NM
Dim C_DOC_CUR
Dim C_OPEN_AMT
Dim C_BAL_AMT
Dim C_CLS_AMT 
Dim C_CLS_LOC_AMT
Dim C_DC_AMT
Dim C_DC_LOC_AMT
Dim C_ITEM_DESC
Dim C_DUE_DT
Dim C_OPEN_NO
Dim C_OPEN_GL_SEQ
Dim C_ORG_CHANGE_ID
Dim C_DEPT_CD
Dim C_DEPT_NM
Dim C_BIZ_AREA_CD
Dim C_BIZ_AREA_NM
Dim C_XCH_RATE

Dim C_ItemSeq		
Dim	C_deptcd
Dim C_deptPopup
Dim C_deptnm
Dim C_AcctCd
Dim C_AcctPopup
Dim C_AcctNm
Dim C_DrCrFg
Dim C_DrCrNm
Dim C_DocCur
Dim C_DocCurPopup
Dim C_ExchRate
Dim C_ItemAmt
Dim C_ItemLocAmt
Dim C_ItemDesc

Dim  lgStrPrevKey1
Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgCurrRow
Dim  lgQueryOk

Dim  intItemCnt					
Dim  IsOpenPop	
Dim  lgRetFlag	                'Popup
Dim  gSelframeFlg

<%
Dim dtToday
dtToday = GetSvrDate
%>

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************

'======================================================================================================
' Name : initSpreadPosVariables()
' Description : 그리드(스프래드) 컬럼 관련 변수 초기화 
'=======================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			C_OPEN_TYPE	     = 1	 
			C_OPEN_TYPE_NM   = 2
			C_GL_NO          = 3
			C_OPEN_DT        = 4
			C_BP_CD          = 5
			C_BP_NM          = 6
			C_OPEN_ACCT_CD   = 7
			C_OPEN_ACCT_NM   = 8
			C_DR_CR_FG       = 9
			C_DR_CR_NM       = 10
			C_DOC_CUR        = 11
			C_OPEN_AMT       = 12
			C_BAL_AMT        = 13
			C_CLS_AMT        = 14
			C_CLS_LOC_AMT    = 15
			C_DC_AMT         = 16
			C_DC_LOC_AMT     = 17			
			C_ITEM_DESC      = 18
			C_DUE_DT         = 19
			C_OPEN_NO        = 20
			C_OPEN_GL_SEQ    = 21
			C_ORG_CHANGE_ID  = 22
			C_DEPT_CD        = 23
			C_DEPT_NM        = 24
			C_BIZ_AREA_CD    = 25
			C_BIZ_AREA_NM	 = 26
			C_XCH_RATE       = 27 	 		
		Case "B"
			C_ItemSeq		 = 1
			C_deptcd         = 2
			C_deptPopup      = 3
			C_deptnm         = 4
			C_AcctCd         = 5
			C_AcctPopup      = 6
			C_AcctNm         = 7
			C_DrCrFg         = 8
			C_DrCrNm         = 9
			C_DocCur         = 10
			C_DocCurPopup    = 11
			C_ExchRate       = 12
			C_ItemAmt        = 13
			C_ItemLocAmt     = 14
			C_ItemDesc       = 15
	End Select			
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = ""                            'initializes Previous Key
    lgStrPrevKey1 = ""
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
    lgSortKey = 1
    lgQueryOk = False
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=======================================================================================================
Sub  SetDefaultVal()
    lgIntFlgMode     = parent.OPMD_CMODE						'Indicates that current mode is Create mode
	frm1.txtAllcDt.text  = UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,gDateFormat)
	frm1.txtDocCur.value = parent.gcurrency
	frm1.hDocCur.value = parent.gcurrency
	frm1.hOrgChangeId.value = parent.gChangeOrgId   	
	lgBlnFlgChgValue     = False
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)

	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspddata4
				ggoSpread.Source = frm1.vspdData4
				ggoSpread.SpreadInit "V20021127",,parent.gAllowDragDropSpread 

				.Redraw = False	

				.MaxCols = C_XCH_RATE + 1   
				.Col = .MaxCols
				.ColHidden = True
				.MaxRows = 0

				Call GetSpreadColumnPos(pvSpdNo)	

				ggoSpread.SSSetEdit  C_OPEN_TYPE     , ""                   , 8 , 3
				ggoSpread.SSSetEdit  C_OPEN_TYPE_NM  , "미결구분"       , 7 , 3				
				ggoSpread.SSSetEdit	 C_GL_NO         , "전표번호"       , 11, 3 
				ggoSpread.SSSetDate	 C_OPEN_DT       , "미결발생일자"   , 11, 2, gDateFormat    
				ggoSpread.SSSetEdit	 C_BP_CD         , "거래처코드"     , 10, 3 
				ggoSpread.SSSetEdit	 C_BP_NM         , "거래처명"       , 13, 3				
				ggoSpread.SSSetEdit	 C_OPEN_ACCT_CD  , "계정코드"       , 7, 3    
				ggoSpread.SSSetEdit	 C_OPEN_ACCT_NM  , "계정명"         , 10, 3
				ggoSpread.SSSetEdit  C_DR_CR_FG      , "차대구분"       , 7 , 3
				ggoSpread.SSSetEdit  C_DR_CR_NM      , "차/대"      , 5	, 3			
				ggoSpread.SSSetEdit	 C_DOC_CUR       , "발생통화"       , 7 , 3
				ggoSpread.SSSetFloat C_OPEN_AMT      , "미결발생금액"   , 12, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_BAL_AMT       , "잔액"           , 12, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_CLS_AMT       , "반제금액"       , 12, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_CLS_LOC_AMT   , "반제금액(자국)" , 12, parent.ggAmTofMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_DC_AMT        , "할인금액"       , 10, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat C_DC_LOC_AMT    , "할인금액(자국)" , 10, parent.ggAmTofMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec				
				ggoSpread.SSSetEdit	 C_ITEM_DESC     , "비고"           , 20, , ,128
				ggoSpread.SSSetDate	 C_DUE_DT        , "만기일자"       , 10, 2, gDateFormat    								
				ggoSpread.SSSetEdit	 C_OPEN_NO       , "미결번호"       , 11, 3 
				ggoSpread.SSSetEdit	 C_OPEN_GL_SEQ   , "미결순번"       , 7, 3 
				ggoSpread.SSSetEdit	 C_ORG_CHANGE_ID , "조직개편아이디" , 5, 3   
				ggoSpread.SSSetEdit	 C_DEPT_CD       , "부서코드"       , 7, 3    
				ggoSpread.SSSetEdit	 C_DEPT_NM       , "부서명"         , 15, 3
				ggoSpread.SSSetEdit	 C_BIZ_AREA_CD   , "사업장"         , 6, 3   
				ggoSpread.SSSetEdit	 C_BIZ_AREA_NM   , "사업장명"       , 20, 3   	 
				ggoSpread.SSSetFloat C_XCH_RATE      , "환율"           , 10, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec				
				
'				Call ggoSpread.MakePairsColumn(C_OPEN_TYPE,C_OPEN_TYPE_NM)
'				Call ggoSpread.MakePairsColumn(C_DR_CR_FG,C_DR_CR_NM)
				Call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
				Call ggoSpread.SSSetColHidden(C_OPEN_TYPE,C_OPEN_TYPE,True)
				Call ggoSpread.SSSetColHidden(C_DR_CR_FG,C_DR_CR_FG,True)
				Call ggoSpread.SSSetColHidden(C_OPEN_GL_SEQ,C_OPEN_GL_SEQ,True)
				
				.Redraw = True 
			End With
		Case "B"
			With frm1.vspddata
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

				.Redraw = False	

				.MaxCols = C_ItemDesc + 1 												'☜: 최대 Columns의 항상 1개 증가시킴 
				.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
				.ColHidden = True       
				.MaxRows = 0		

				Call GetSpreadColumnPos(pvSpdNo)    
				Call AppendNumberPlace("6","3","0")

				ggoSpread.SSSetFloat  C_ItemSeq     , "NO"			, 6,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"
				ggoSpread.SSSetEdit	  C_deptcd      , "부서코드"    , 10, , , 10, 2
				ggoSpread.SSSetButton C_deptPopup
				ggoSpread.SSSetEdit	  C_deptnm      , "부서명"      , 17, 3
				ggoSpread.SSSetEdit	  C_AcctCd      , "계정코드"    , 15, , , 18
				ggoSpread.SSSetButton C_AcctPopup
				ggoSpread.SSSetEdit	  C_AcctNm      , "계정명"      , 20, , , 30
				ggoSpread.SSSetCombo  C_DrCrFg      , "차대구분"    , 8
				ggoSpread.SSSetCombo  C_DrCrNm      , "차대구분"    , 10
				ggoSpread.SSSetEdit	  C_DocCur      , "거래통화"    , 10, , , 10, 2
				ggoSpread.SSSetButton C_DocCurPopup
				ggoSpread.SSSetFloat  C_ExchRate    , "환율"        , 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_ItemAmt     , "금액"        , 15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetFloat  C_ItemLocAmt  , "금액(자국)"	, 15, parent.ggAmTofMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
				ggoSpread.SSSetEdit	  C_ItemDesc    , "비고"        , 30, , ,128

				Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)
				Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)				
				
				.Redraw = True    
			End With
	End Select
	
    Call SetSpreadLock(pvSpdNo)
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock(ByVal pvSpdNo)														'Form_Load, Query후 그리드 세팅 

	Dim ii

	With frm1
		Select Case UCase(Trim(pvSpdNo))
			Case "A"
				ggoSpread.Source = .vspddata4
				.vspddata4.ReDraw = False

				ggoSpread.SpreadLock    C_OPEN_TYPE     ,-1, C_OPEN_TYPE      , -1
				ggoSpread.SpreadLock    C_OPEN_TYPE_NM  ,-1, C_OPEN_TYPE_NM   , -1
				ggoSpread.SpreadLock    C_GL_NO	        ,-1, C_GL_NO          , -1				
				ggoSpread.SpreadLock    C_OPEN_DT       ,-1, C_OPEN_DT        , -1
				ggoSpread.SpreadLock    C_BP_CD         ,-1, C_BP_CD          , -1
				ggoSpread.SpreadLock    C_BP_NM	        ,-1, C_BP_NM          , -1			
				ggoSpread.SpreadLock    C_OPEN_ACCT_CD  ,-1, C_OPEN_ACCT_CD   , -1
				ggoSpread.SpreadLock    C_OPEN_ACCT_NM  ,-1, C_OPEN_ACCT_NM   , -1
				ggoSpread.SpreadLock    C_DR_CR_FG      ,-1, C_DR_CR_FG       , -1
				ggoSpread.SpreadLock    C_DR_CR_NM      ,-1, C_DR_CR_NM       , -1
				ggoSpread.SpreadLock    C_DOC_CUR       ,-1, C_DOC_CUR        , -1
				ggoSpread.SpreadLock    C_OPEN_AMT      ,-1, C_OPEN_AMT       , -1
				ggoSpread.SpreadLock    C_BAL_AMT       ,-1, C_BAL_AMT        , -1
				ggoSpread.SpreadUnLock  C_CLS_AMT       ,-1, C_CLS_AMT        , -1																			
				ggoSpread.SSSetRequired C_CLS_AMT       ,-1,  -1
				ggoSpread.SpreadUnLock  C_CLS_LOC_AMT   ,-1, C_CLS_LOC_AMT    , -1
				ggoSpread.SpreadUnLock  C_DC_AMT        ,-1, C_DC_AMT         , -1																		
				ggoSpread.SpreadUnLock  C_DC_LOC_AMT    ,-1, C_DC_LOC_AMT     , -1
				ggoSpread.SpreadUnLock  C_ITEM_DESC     ,-1, C_ITEM_DESC      , -1
				ggoSpread.SpreadLock    C_DUE_DT        ,-1, C_DUE_DT         , -1
				ggoSpread.SpreadLock    C_OPEN_NO       ,-1, C_OPEN_NO        , -1
				ggoSpread.SpreadLock    C_OPEN_GL_SEQ   ,-1, C_OPEN_GL_SEQ    , -1
				ggoSpread.SpreadLock    C_ORG_CHANGE_ID ,-1, C_ORG_CHANGE_ID  , -1
				ggoSpread.SpreadLock    C_DEPT_CD       ,-1, C_DEPT_CD        , -1
				ggoSpread.SpreadLock    C_DEPT_NM       ,-1, C_DEPT_NM        , -1
				ggoSpread.SpreadLock    C_BIZ_AREA_CD   ,-1, C_BIZ_AREA_CD    , -1
				ggoSpread.SpreadLock    C_BIZ_AREA_NM   ,-1, C_BIZ_AREA_NM    , -1
				ggoSpread.SpreadLock    C_XCH_RATE      ,-1, C_XCH_RATE    , -1				
				
				For ii = 1 To .vspddata4.MaxRows
					.vspddata4.Col = C_OPEN_TYPE
					.vspddata4.Row = ii
					If Trim(.vspddata4.Text) = "AR" Or Trim(.vspddata4.Text) = "AP" Then

					Else
						ggoSpread.SSSetProtected  C_DC_AMT        ,ii,  ii																		
						ggoSpread.SSSetProtected  C_DC_LOC_AMT    ,ii,  ii		
					End If
				Next

				.vspddata4.ReDraw = True   
			Case "B"	
				ggoSpread.Source = .vspddata
				.vspddata.Redraw = False    

				ggoSpread.SpreadLock    C_ItemSeq      ,-1, C_ItemSeq       , -1
				ggoSpread.SpreadUnLock  C_deptcd       ,-1, C_deptcd        , -1
				ggoSpread.SSSetRequired C_deptcd       ,-1,  -1
				ggoSpread.SpreadUnLock  C_deptPopup    ,-1, C_deptPopup     , -1
				ggoSpread.SpreadLock    C_deptnm       ,-1, C_deptnm        , -1
				ggoSpread.SpreadLock	C_AcctCd       ,-1, C_AcctCd        , -1
				ggoSpread.SpreadLock	C_AcctPopup    ,-1, C_AcctPopup     , -1
				ggoSpread.SpreadLock    C_AcctNm       ,-1, C_AcctNm        , -1
				ggoSpread.SpreadUnLock  C_DrCrFg       ,-1, C_DrCrFg        , -1
				ggoSpread.SSSetRequired C_DrCrFg       ,-1,  -1
				ggoSpread.SpreadUnLock  C_DrCrNm       ,-1, C_DrCrNm        , -1
				ggoSpread.SSSetRequired C_DrCrNm       ,-1,  -1
				ggoSpread.SpreadUnLock  C_DocCur       ,-1, C_DocCur        , -1
				ggoSpread.SSSetRequired C_DocCur       ,-1,  -1
				ggoSpread.SpreadUnLock  C_DocCurPopup  ,-1, C_DocCurPopup   , -1
				ggoSpread.SpreadUnLock  C_ExchRate     ,-1, C_ExchRate      , -1
				ggoSpread.SpreadUnLock  C_ItemAmt      ,-1, C_ItemAmt       , -1
				ggoSpread.SSSetRequired C_ItemAmt      ,-1,  -1
				ggoSpread.SpreadUnLock  C_ItemLocAmt   ,-1, C_ItemLocAmt    , -1
				ggoSpread.SpreadUnLock  C_ItemDesc     ,-1, C_ItemDesc      , -1

				.vspddata.ReDraw = True   		
		End Select			
		
	
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColorOpen
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColorOpen(ByVal pvStartRow , ByVal pvEndRow)							'행추가, 행복사 후 추가된 그리드 세팅 
	Dim ii

	With frm1
		ggoSpread.Source = .vspddata4
		.vspddata4.ReDraw = False

		ggoSpread.SpreadLock    C_OPEN_TYPE     ,pvStartRow, C_OPEN_TYPE      , pvEndRow
		ggoSpread.SpreadLock    C_OPEN_TYPE_NM  ,pvStartRow, C_OPEN_TYPE_NM   , pvEndRow
		ggoSpread.SpreadLock    C_GL_NO	        ,pvStartRow, C_GL_NO          , pvEndRow				
		ggoSpread.SpreadLock    C_OPEN_DT       ,pvStartRow, C_OPEN_DT        , pvEndRow
		ggoSpread.SpreadLock    C_BP_CD         ,pvStartRow, C_BP_CD          , pvEndRow
		ggoSpread.SpreadLock    C_BP_NM	        ,pvStartRow, C_BP_NM          , pvEndRow			
		ggoSpread.SpreadLock    C_OPEN_ACCT_CD  ,pvStartRow, C_OPEN_ACCT_CD   , pvEndRow
		ggoSpread.SpreadLock    C_OPEN_ACCT_NM  ,pvStartRow, C_OPEN_ACCT_NM   , pvEndRow
		ggoSpread.SpreadLock    C_DR_CR_FG      ,pvStartRow, C_DR_CR_FG       , pvEndRow
		ggoSpread.SpreadLock    C_DR_CR_NM      ,pvStartRow, C_DR_CR_NM       , pvEndRow
		ggoSpread.SpreadLock    C_DOC_CUR       ,pvStartRow, C_DOC_CUR        , pvEndRow
		ggoSpread.SpreadLock    C_OPEN_AMT      ,pvStartRow, C_OPEN_AMT       , pvEndRow
		ggoSpread.SpreadLock    C_BAL_AMT       ,pvStartRow, C_BAL_AMT        , pvEndRow
		ggoSpread.SpreadUnLock  C_CLS_AMT       ,pvStartRow, C_CLS_AMT        , pvEndRow																			
		ggoSpread.SSSetRequired C_CLS_AMT       ,pvStartRow,  pvEndRow
		ggoSpread.SpreadUnLock  C_CLS_LOC_AMT   ,pvStartRow, C_CLS_LOC_AMT    , pvEndRow
		ggoSpread.SpreadUnLock  C_DC_AMT        ,pvStartRow, C_DC_AMT         , pvEndRow																		
		ggoSpread.SpreadUnLock  C_DC_LOC_AMT    ,pvStartRow, C_DC_LOC_AMT     , pvEndRow
		ggoSpread.SpreadUnLock  C_ITEM_DESC     ,pvStartRow, C_ITEM_DESC      , pvEndRow
		ggoSpread.SpreadLock    C_DUE_DT        ,pvStartRow, C_DUE_DT         , pvEndRow
		ggoSpread.SpreadLock    C_OPEN_NO       ,pvStartRow, C_OPEN_NO        , pvEndRow
		ggoSpread.SpreadLock    C_ORG_CHANGE_ID ,pvStartRow, C_ORG_CHANGE_ID  , pvEndRow
		ggoSpread.SpreadLock    C_DEPT_CD       ,pvStartRow, C_DEPT_CD        , pvEndRow
		ggoSpread.SpreadLock    C_DEPT_NM       ,pvStartRow, C_DEPT_NM        , pvEndRow
		ggoSpread.SpreadLock    C_BIZ_AREA_CD   ,pvStartRow, C_BIZ_AREA_CD    , pvEndRow
		ggoSpread.SpreadLock    C_BIZ_AREA_NM   ,pvStartRow, C_BIZ_AREA_NM    , pvEndRow
		ggoSpread.SpreadLock    C_XCH_RATE      ,pvStartRow, C_XCH_RATE    , pvEndRow		
		

		.vspddata4.ReDraw = True
    End With
End Sub


'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColorAcct(ByVal pvStartRow , ByVal pvEndRow)							'행추가, 행복사 후 추가된 그리드 세팅 
	With frm1.vspdData
		.Redraw = False
		ggoSpread.Source = frm1.vspdData			
		ggoSpread.SSSetProtected C_ItemSeq   , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_deptcd   , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_deptnm	  , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_AcctCd    , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_AcctNm	  , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DrCrNm  , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_DocCur   , pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_ItemAmt   , pvStartRow, pvEndRow		

		.Col = 2																	'컬럼의 절대 위치로 이동 
		.Row = .ActiveRow
		.Action = 0                         
		.EditMode = True		
		.Redraw = True		
    End With		
End Sub

'======================================================================================================
' Function Name : SetSpread2ColorCtrl
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub  SetSpread2ColorCtrl(ByVal Row)
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False	 
	
		For i = 1 To .vspdData2.MaxRows
			ggoSpread.SSSetProtected C_DtlSeq   , i, i
			ggoSpread.SSSetProtected C_CtrlCd   , i, i
			ggoSpread.SSSetProtected C_CtrlNm   , i, i			
			ggoSpread.SSSetProtected C_CtrlValNm, i, i
			
			.vspdData.Row = Row
			.vspdData.Col = C_DrCrFg						
			
			If Trim(.vspddata.Text) = "DR" Then
				.vspdData2.Row = i
				.vspdData2.Col = C_DrFg

				If (.vspdData2.text = "Y")  Or (.vspdData2.text = "DC") Or (.vspdData2.text = "D") Then
					ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
				End If
			Elseif Trim(.vspddata.Text) = "CR" Then
				.vspdData2.Row = i
				.vspdData2.Col = C_DrFg

				If (.vspdData2.text = "Y")  Or (.vspdData2.text = "DC") Or (.vspdData2.text = "C") Then
					ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
				End If
			Else
				.vspdData2.Row = i
				.vspdData2.Col = C_DrFg

				If (.vspdData2.text = "Y")  Or (.vspdData2.text = "DC")  Then
					ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
				End If			
			
			End If				
		Next

		.vspdData2.ReDraw = True
    End With
End Sub

'============================================================================================================
Function InitComboBoxGrid()
    ggoSpread.Source = frm1.vspdData

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm
End Function



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
			
			C_OPEN_TYPE	     = iCurColumnPos(1)
			C_OPEN_TYPE_NM   = iCurColumnPos(2)
			C_GL_NO          = iCurColumnPos(3)
			C_OPEN_DT        = iCurColumnPos(4)
			C_BP_CD          = iCurColumnPos(5)
			C_BP_NM          = iCurColumnPos(6)
			C_OPEN_ACCT_CD   = iCurColumnPos(7)
			C_OPEN_ACCT_NM   = iCurColumnPos(8)
			C_DR_CR_FG       = iCurColumnPos(9)
			C_DR_CR_NM       = iCurColumnPos(10)
			C_DOC_CUR        = iCurColumnPos(11)
			C_OPEN_AMT       = iCurColumnPos(12)
			C_BAL_AMT        = iCurColumnPos(13)
			C_CLS_AMT        = iCurColumnPos(14)
			C_CLS_LOC_AMT    = iCurColumnPos(15)
			C_DC_AMT         = iCurColumnPos(16)
			C_DC_LOC_AMT     = iCurColumnPos(17)			
			C_ITEM_DESC      = iCurColumnPos(18)
			C_DUE_DT         = iCurColumnPos(19)
			C_OPEN_NO        = iCurColumnPos(20)
			C_OPEN_GL_SEQ    = iCurColumnPos(21)
			C_ORG_CHANGE_ID  = iCurColumnPos(22)
			C_DEPT_CD        = iCurColumnPos(23)
			C_DEPT_NM        = iCurColumnPos(24)
			C_BIZ_AREA_CD    = iCurColumnPos(25)
			C_BIZ_AREA_NM	 = iCurColumnPos(26)
			C_XCH_RATE       = iCurColumnPos(27)
		Case "B"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)							

			C_ItemSeq		 = iCurColumnPos(1)
			C_deptcd         = iCurColumnPos(2)
			C_deptPopup      = iCurColumnPos(3)
			C_deptnm         = iCurColumnPos(4)
			C_AcctCd         = iCurColumnPos(5)
			C_AcctPopup      = iCurColumnPos(6)
			C_AcctNm         = iCurColumnPos(7)
			C_DrCrFg         = iCurColumnPos(8)
			C_DrCrNm         = iCurColumnPos(9)
			C_DocCur         = iCurColumnPos(10)
			C_DocCurPopup    = iCurColumnPos(11)
			C_ExchRate       = iCurColumnPos(12)
			C_ItemAmt        = iCurColumnPos(13)
			C_ItemLocAmt     = iCurColumnPos(14)
			C_ItemDesc       = iCurColumnPos(15)
	End select
End Sub

'======================================================================================================
'	Name : Open???()
'	Description : Ref 화면을 call한다. 
'======================================================================================================
Function OpenRefOpenNo()
	Dim arrRet
	Dim arrParam(14)
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5150ra2_ko441")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5150ra2_ko441", "X")
		IsOpenPop = False
		Exit Function
	End If

	If gSelframeFlg <> TAB1 Then Exit Function		 		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
		If .vspddata4.MaxRows = 0 Then .hDocCur.value	= ""

		arrParam(0) = .txtBpCd.value											' 검색조건이 있을경우 파라미터 
		arrParam(1) = .txtBpNm.value			
		arrParam(2) = .txtDocCur.value
		arrParam(3) = "M"
		arrParam(6) = .txtAllcDt.text
		arrParam(7) = .txtAllcDt.Alt
		arrParam(8) = .hDocCur.value
		arrParam(9) = .txtGlNo.value
	End With

	' 권한관리 추가 
	arrParam(11) = lgAuthBizAreaCd
	arrParam(12) = lgInternalCd
	arrParam(13) = lgSubInternalCd
	arrParam(14) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=960px; dialogHeight=600px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0, 0) = "" Then		
		Exit Function
	Else		
		Call SetRefOpen(arrRet)
	End If
End Function

'======================================================================================================
'	Name : SetRefOpenAr()
'	Description : OpenAp Popup에서 Return되는 값 setting
'======================================================================================================
Function SetRefOpen(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	DIM X
	Dim sFindFg
	Dim ii
	Dim iOpenNo , iGLSeq
	
	With frm1
		.vspddata4.focus		
		ggoSpread.Source = .vspddata4
		.vspddata4.ReDraw = False	
	
		TempRow = .vspddata4.MaxRows												'☜: 현재까지의 MaxRows

		For I = TempRow To TempRow + Ubound(arrRet, 1)
			sFindFg	= "N"
			For x = 1 To TempRow
				.vspddata4.Row = x
				.vspddata4.Col = C_OPEN_NO
				iOpenNo = .vspddata4.text
				
				.vspddata4.Row = x
				.vspddata4.Col = C_OPEN_GL_SEQ
				iGLSeq = .vspddata4.text 								
				
				If UCase(Trim(iOpenNo)) = UCase(Trim(arrRet(I - TempRow, 14)))  And  Trim(iGLSeq) = Trim(arrRet(I - TempRow, 22))  Then
					sFindFg	= "Y"
				End If
			Next
			
			If 	sFindFg	= "N" Then
				.vspddata4.MaxRows = .vspddata4.MaxRows + 1
				.vspddata4.Row = I + 1				
				.vspddata4.Col = 0
				.vspddata4.Text = ggoSpread.InsertFlag

				.vspddata4.Col = C_OPEN_TYPE
				.vspddata4.text = arrRet(I - TempRow,12)
				.vspddata4.Col = C_OPEN_TYPE_NM
				.vspddata4.text = arrRet(I - TempRow,13)
				.vspddata4.Col = C_GL_NO
				.vspddata4.text = arrRet(I - TempRow,0)
				.vspddata4.Col = C_OPEN_DT
				.vspddata4.text = arrRet(I - TempRow,1)
				.vspddata4.Col = C_BP_CD
				.vspddata4.text = arrRet(I - TempRow,2)
				.vspddata4.Col = C_BP_NM
				.vspddata4.text = arrRet(I - TempRow,3)
				.vspddata4.Col = C_OPEN_ACCT_CD
				.vspddata4.text = arrRet(I - TempRow,4)
				.vspddata4.Col = C_OPEN_ACCT_NM
				.vspddata4.text = arrRet(I - TempRow,5)
				.vspddata4.Col = C_DR_CR_FG
				.vspddata4.text = arrRet(I - TempRow,6)
				.vspddata4.Col = C_DR_CR_NM
				.vspddata4.text = arrRet(I - TempRow,7)
				.vspddata4.Col = C_DOC_CUR
				.vspddata4.text = arrRet(I - TempRow,8)
				.vspddata4.Col = C_OPEN_AMT
				.vspddata4.text = arrRet(I - TempRow,10)
				.vspddata4.Col = C_BAL_AMT
				.vspddata4.text = arrRet(I - TempRow,9)
				.vspddata4.Col = C_CLS_AMT
				.vspddata4.text = arrRet(I - TempRow,21)
				.vspddata4.Col = C_CLS_LOC_AMT
				.vspddata4.text = ""
				.vspddata4.Col = C_DC_AMT
				.vspddata4.text = ""
				.vspddata4.Col = C_DC_LOC_AMT
				.vspddata4.text = ""
				.vspddata4.Col = C_ITEM_DESC
				.vspddata4.text = arrRet(I - TempRow,24)
				.vspddata4.Col = C_DUE_DT
				.vspddata4.text = arrRet(I - TempRow,15)				
				.vspddata4.Col = C_OPEN_NO
				.vspddata4.text = arrRet(I - TempRow,14)
				.vspddata4.Col = C_OPEN_GL_SEQ
				.vspddata4.text = arrRet(I - TempRow,22)				
				.vspddata4.Col = C_ORG_CHANGE_ID
				.vspddata4.text = arrRet(I - TempRow,16)
				.vspddata4.Col = C_DEPT_CD
				.vspddata4.text = arrRet(I - TempRow,17)
				.vspddata4.Col = C_DEPT_NM
				.vspddata4.text = arrRet(I - TempRow,18)
				.vspddata4.Col = C_BIZ_AREA_CD
				.vspddata4.text = arrRet(I - TempRow,19)
				.vspddata4.Col = C_BIZ_AREA_NM
				.vspddata4.text = arrRet(I - TempRow,20)
				.vspddata4.Col = C_XCH_RATE
				.vspddata4.text = arrRet(I - TempRow,23)


				.vspddata4.Col = C_OPEN_TYPE
				If .vspddata4.text <> "U6" Then
					.txtBpCd.value = arrRet(I - TempRow,2)
					.txtBpNm.value = arrRet(I - TempRow,3)
				End If		
			End If	
		Next

		If arrRet(0, 8) <> parent.gCurrency Then
			.hDocCur.Value = arrRet(0, 8)
			.txtDocCur.Value = arrRet(0, 8)
			.hDocCur2.value = arrRet(0, 8)
			Call ggoOper.SetReqAttr(.txtDocCur,"Q")
		End If

		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,TempRow + 1,.vspddata4.MaxRows,C_DOC_CUR,C_OPEN_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,TempRow + 1,.vspddata4.MaxRows,C_DOC_CUR,C_BAL_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,TempRow + 1,.vspddata4.MaxRows,C_DOC_CUR,C_CLS_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,TempRow + 1,.vspddata4.MaxRows,C_DOC_CUR,C_XCH_RATE,"D", "I" ,"X","X")

		If TempRow + 1 <= .vspddata4.MaxRows Then
			Call SetSpreadColorOpen(TempRow + 1,.vspddata4.MaxRows)

			For ii = TempRow + 1 To .vspddata4.MaxRows
				.vspddata4.Col = C_OPEN_TYPE
				.vspddata4.Row = ii
				If Trim(.vspddata4.Text) = "AR" Or Trim(.vspddata4.Text) = "AP" Then

				Else
					ggoSpread.SSSetProtected  C_DC_AMT        ,ii,  ii																		
					ggoSpread.SSSetProtected  C_DC_LOC_AMT    ,ii,  ii		
				End If
			Next
		End If

		.vspddata4.ReDraw = True
    End With
    
    Call DoSum()
End Function

'======================================================================================================
'	기능: 
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function MoveJmpClick()
	Select Case gSelframeFlg
		Case TAB1
			spnRef.innerHTML =  "<a href='vbscript:OpenRefOpenNo()'>미결통합참조</A>&nbsp;|&nbsp;"
		Case TAB2
			spnRef.innerHTML = "<font color=""#777777"">미결통합참조</font>&nbsp;|&nbsp;"
	End Select    
End Function

'======================================================================================================
'   Function Name : OpenPopUpgl()
'   Function Desc : 
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
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'회계전표번호 
	arrParam(1) = ""						'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

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
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)					'결의전표번호 
	arrParam(1) = ""											'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iArrParam(8)
	Dim strCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
'MsgBox iWhere
	Select Case iWhere
		Case 0
		
		Case 1
			arrParam(0) = "통화코드 POPUP"								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"										' TABLE 명칭 
			frm1.vspddata.Row = frm1.vspddata.ActiveRow
			frm1.vspddata.Col = C_DocCur
			arrParam(2) = Trim(frm1.vspddata.Text)							' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "통화코드"			
	
			arrField(0) = "CURRENCY"										' Field명(0)
			arrField(1) = "CURRENCY_DESC"									' Field명(1)
    
			arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드설명"
		Case 2		
			arrParam(0) = "통화코드 POPUP"								' 팝업 명칭 
			arrParam(1) = "B_CURRENCY"										' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtDocCur.Value)						' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""												' Where Condition
			arrParam(5) = "통화코드"			
	
			arrField(0) = "CURRENCY"										' Field명(0)
			arrField(1) = "CURRENCY_DESC"									' Field명(1)
    
			arrHeader(0) = "통화코드"									' Header명(0)
			arrHeader(1) = "통화코드설명"
		Case 3
			arrParam(0) = "계정코드 POPUP"								' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE 명칭 
			arrParam(2) = Trim(strCode)										' Code Condition
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
	End Select				
		
	If iwhere = 0 Then		
		Dim iCalledAspName
		iCalledAspName = AskPRAspName("a5150ra1")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5150ra1", "X")
			IsOpenPop = False
			Exit Function
		End If			
		
		' 권한관리 추가 
		iArrParam(5) = lgAuthBizAreaCd
		iArrParam(6) = lgInternalCd
		iArrParam(7) = lgSubInternalCd
		iArrParam(8) = lgAuthUsrID		

		arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,iArrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	else
		'if iWhere = 1 or iWhere = 3 Then
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		'else
		'	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		'		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		'End if
	End If
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

'======================================================================================================
'   Function Name : EscPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		
				.txtAllcNo.focus		
			Case 1
				Call SetActiveCell(frm1.vspdData,C_DocCur,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 2
				.txtDocCur.focus
			Case 3
				Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
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
				.txtAllcNo.value = arrRet(0)
				.txtAllcNo.focus
			Case 1
				.vspddata.Col = C_DocCur
				.vspddata.Text = arrRet(0)
			
				Call vspddata_Change(C_DocCur, frm1.vspddata.activerow )	 ' 변경이 일어났다고 알려줌 
				Call SetActiveCell(frm1.vspdData,C_DocCur,frm1.vspdData.ActiveRow ,"M","X","X")
			Case 2
				.txtDocCur.value = arrRet(0)		
				.txtDocCur.focus
			Case 3
				.vspddata.Col = C_AcctCd
				.vspddata.Text = arrRet(0)
				.vspddata.Col = C_AcctNm
				.vspddata.Text = arrRet(1)

				Call vspddata_Change(C_AcctCd, frm1.vspddata.activerow )	 ' 변경이 일어났다고 알려줌 
				Call SetActiveCell(frm1.vspdData,C_AcctCd,frm1.vspdData.ActiveRow ,"M","X","X")
		End Select
	End With

	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End if	
End Function

 '------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Bp PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd(ByVal strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처팝업"
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "거래처"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    
    arrHeader(0) = "거래처"		
    arrHeader(1) = "거래처명"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	IF 	arrRet(0) <> "" then			
		Call SetBpCd(arrRet)
	Else
		frm1.txtBpCd.focus
	end if
End Function

 '------------------------------------------  SetBpCd()  --------------------------------------------------
'	Name : SetBpCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBpCd(Byval arrRet)
	frm1.txtBpCd.value = arrRet(0)
	frm1.txtBpNm.value = arrRet(1)
	frm1.txtBpCd.focus
	lgBlnFlgChgValue = True
End Function

'======================================================================================================
'	Name : OpenDept
'	Description : 
'=======================================================================================================%>
Function OpenDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = frm1.txtAllcDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = "F"									' 결의일자 상태 Condition  
	
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID		

	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = "0" Then
			frm1.txtDeptCd.focus
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
		Select Case iWhere
		    Case "0"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
				.txtAllcDt.text = arrRet(3)
				Call txtDeptCd_OnBlur()  
				frm1.txtDeptCd.focus
			Case "1"
				.vspddata.Col = C_deptcd
				.vspddata.Row = .vspddata.ActiveRow
				.vspddata.Text = arrRet(0)
				.vspddata.Col = C_deptnm
				.vspddata.Row = .vspddata.ActiveRow
				.vspddata.Text = arrRet(1)				
	    End Select
	End With
	
  	lgBlnFlgChgValue = True	
End Function 

'==========================================================================================
'   Event Name : txtDeptCd_OnBlur
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnBlur()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtAllcDt.Text = "") Then    
		Exit sub
    End If
    
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(frm1.txtDeptCd.value, "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			

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
	End If	
End Sub

'==========================================================================================
'   Event Name : txtAllcDt_onBlur
'   Event Desc : 
'==========================================================================================
Sub txtAllcDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
  	lgBlnFlgChgValue = True

	With frm1
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtAllcDt.Text <> "") Then
			strSelect	=			 " Distinct org_change_id "    		
			strFrom		=			 " b_acct_dept(NOLOCK) "		
			strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
			strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtAllcDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
				.txtDeptCd.value = ""
				.txtDeptNm.value = ""
				.hOrgChangeId.value = ""
				.txtDeptCd.focus
			End If
		End If
	End With
End Sub

'========================================================================================================= 
Function OpenUnderDept(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim field_fg

	IsOpenPop = True

	'If RTrim(LTrim(frm1.txtDeptCd.value)) <> "" Then
	'	arrParam(0) = "부서 팝업"	
	'	arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"
	'	arrParam(2) = Trim(strCode)
	'	arrParam(3) = "" 
	'	arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ""
	'	arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And B.BIZ_AREA_CD = ( Select B.BIZ_AREA_CD"
	'	arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.value , "''", "S") & ""
	'	arrParam(4) = arrParam(4) & " And A.COST_CD = B.COST_CD And A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.value , "''", "S") & ")"
	'	arrParam(5) = "부서코드"

	'	arrField(0) = "A.DEPT_CD"
	'	arrField(1) = "A.DEPT_Nm"
	'	arrField(2) = "B.BIZ_AREA_CD"

	'	arrHeader(0) = "부서코드"
	'	arrHeader(1) = "부서코드명"
	'	arrHeader(2) = "사업장코드"
		
'		arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=B_ACCT_DEPT_01", _
'		   Array(Array(Trim(strCode)),Array("3",Trim(frm1.hOrgChangeId.value),Trim(frm1.txtDeptCd.value),Trim(frm1.hOrgChangeId.value))), _
'		   "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	 		
	'Else
		arrParam(0) = "부서 팝업"	
		arrParam(1) = "B_ACCT_DEPT A"
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = (Select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = ( Select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtAllcDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		arrParam(5) = "부서코드"
		arrField(0) = "A.DEPT_CD"
		arrField(1) = "A.DEPT_Nm"
		arrHeader(0) = "부서코드"
		arrHeader(1) = "부서코드명"

'		arrRet = window.showModalDialog("../../comasp/CommonPopup2.asp?pid=B_ACCT_DEPT_00", Array(Array(Trim(strCode))), _
'		           "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")  	 
	'End If

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetDept(arrRet, iWhere)
	End If

	Call FocusAfterDeptPopup(iWhere)
End Function


'=======================================================================================================
Function FocusAfterDeptPopup(ByVal iWhere)
	With frm1
		Select Case iWhere
			Case 0  
				.txtDeptCd.focus
			Case 1 
				Call SetActiveCell(.vspdData,C_deptcd,.vspdData.ActiveRow ,"M","X","X")
		End Select    
	End With
End Function

'======================================================================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'=======================================================================================================
Function ClickTab1()
	
	If gSelframeFlg = TAB1 Then Exit Function

	Call changeTabs(TAB1)	 '~~~ 첫번째 Tab 

	gSelframeFlg = TAB1	

	Call MoveJmpClick

	If lgIntFlgMode <> parent.OPMD_UMODE Then
	    Call SetToolBar("1111101100001111")										'⊙: 버튼 툴바 제어 
	Else                 
	    Call SetToolBar("1111101100001111")										'⊙: 버튼 툴바 제어 
	End If
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)	 '~~~ 첫번째 Tab 

	gSelframeFlg = TAB2

	Call MoveJmpClick

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetToolBar("1110111100001111")
	Else                 
		Call SetToolBar("1111111100001111")
	End If	
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************




'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub  Form_Load()
    Call LoadInfTB19029()																	'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)							
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")													'Lock  Suitable  Field
    Call InitSpreadSheet("A")																'Setup the Spread sheet
    Call InitSpreadSheet("B")																'Setup the Spread sheet    
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	    
    Call InitVariables()																	'Initializes local global variables
    Call SetDefaultVal()
	Call InitComboBoxGrid()
	Call ClickTab1()

	frm1.txtAllcNo.focus

	gIsTab     = "Y" 
	gTabMaxCnt = 2  

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

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related To Query ButTon of Main ToolBar
'=======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    Dim var1, var2,var3

    FncQuery = False                                                        

    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then													'This function check indispensable field
		Exit Function
    End If

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspddata4
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata2
    var3 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True  Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()																'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspddata4
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																		'☜: Query db data

    FncQuery = True		

	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related To New ButTon of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1, var2, var3
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspddata4
    var1 = ggoSpread.SSCheckChange
	ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange
    
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Or var1 = True Or var2 = True Or var3 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")										'Clear Condition Field
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
    Call InitVariables()																'Initializes local global variables
    Call SetDefaultVal()    
    Call DisableRefPop()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspddata4
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus

    lgBlnFlgChgValue = False            
    FncNew = True                                                          

	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related To Delete ButTon of Main ToolBar
'========================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then											'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")						'Will you desTory previous data"

	If IntRetCD = vbNo Then
		Exit Function
	End If
	
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then															'☜: Delete db data
		Exit Function																	'☜:
    End If					
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------    
    FncDelete = True                                                        
		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related To Delete ButTon of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2, var3
	
    FncSave = False                                                         
    
    Err.Clear                                                               
	        
    ggoSpread.Source = frm1.vspddata4
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspddata2
    var3 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False And var1 = False And var2 = False And var3 = False Then	'⊙: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")									'⊙: Display Message(There is no changed data.)
		Exit Function		
    End If

	If CheckSpread3 = False Then
		IntRetCD = DisplayMsgBox ("110420", "X", "X", "X")
	    Exit Function
	End If

    
    If Not chkField(Document, "2") Then													'⊙: Check required field(Single area)
		Exit Function
    End If    
    
    ggoSpread.Source = frm1.vspddata4
    If Not ggoSpread.SSDefaultCheck Then
		Call ClickTab()											'⊙: Check contents area
		Exit Function
    End If

	ggoSpread.Source = frm1.vspddata
    If Not ggoSpread.SSDefaultCheck Then											'⊙: Check contents area
		Call ClickTab2()
		Exit Function
    End If
    
    If Not chkAllcDate Then
		Exit Function
    End If  
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																		'☜: Save db data
    
    FncSave = True                                                       
		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related To Copy ButTon of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	
	If frm1.vspddata.Maxrows < 1 Then Exit Function  
	
	frm1.vspddata.ReDraw = False
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")	'⊙: "Will you desTory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	
	With frm1
		.vspddata.ReDraw = False
	
		ggoSpread.Source = .vspddata	
		ggoSpread.CopyRow

		Call SetSpreadColorAcct(.vspddata.ActiveRow, .vspddata.ActiveRow)

		.vspddata.ReDraw = True
	End With
			
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related To Cancel ButTon of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	With frm1
		If gSelframeFlg = TAB1 Then
			If .vspddata4.Maxrows < 1 Then Exit Function

			.vspddata4.Row = .vspddata4.ActiveRow
			.vspddata4.Col = 0

			ggoSpread.Source = .vspddata4
			ggoSpread.EditUndo

			Call Dosum()

			If .vspddata4.MaxRows < 1 Then 
				Call ggoOper.SetReqAttr(.txtAllcDt,   "N")
				Exit Function
			End If					
		Else
			If .vspdData.MaxRows < 1 Then Exit Function

		    .vspdData.Row = .vspdData.ActiveRow
		    .vspdData.Col = 0

		    If .vspdData.Text = ggoSpread.InsertFlag Then
				.vspdData.Col = C_AcctCd
				If Len(Trim(.vspdData.Text)) > 0 Then 
					.vspdData.Col = C_ItemSeq
					DeleteHSheet(.vspdData.Text)
				End If		
		    End If

		    ggoSpread.Source = .vspdData	
		    ggoSpread.EditUndo

			Call DoSum()

			If .vspdData.MaxRows < 1 Then Exit Function
			
			.vspdData.Row = .vspdData.ActiveRow
			.vspdData.Col = 0		    
			
			If .vspdData.Row = 0 Then Exit Function

		    If .vspdData.Text = ggoSpread.InsertFlag Then
				.vspdData.Col = C_AcctCd
				If Len(Trim(.text)) > 0 Then
					.vspdData.Col = C_ItemSeq
					.hItemSeq.value = .vspdData.Text
					ggoSpread.Source = .vspdData2
					ggoSpread.ClearSpreadData

					Call DbQuery3(.vspdData.ActiveRow)
				End If
		    Else
		        .vspdData.Col = C_ItemSeq
		        .hItemSeq.value = .vspdData.Text
				ggoSpread.Source = .vspdData2
				ggoSpread.ClearSpreadData
		        Call DbQuery2(.vspdData.ActiveRow)
		    End If
		End If
	End With

	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related To InsertRow ButTon of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowCnt) 
	Dim imRow
	Dim ii
	Dim iCurRowPos
	
    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error stat	

'	If Not chkFieldByCell(.txtDeptCd, "A", "1") Then Exit Function

    If gSelframeFlg <> TAB2 Then
		Call ClickTab2()															'sstData.Tab = 1
    End If

	FncInsertRow = False															'☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End If		

	With frm1
		iCurRowPos = .vspdData.ActiveRow
        .vspdData.ReDraw = False
        ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		For ii = .vspdData.ActiveRow To  .vspdData.ActiveRow + imRow - 1
			.vspddata.row = ii
			.vspddata.col = C_DocCur
			.vspddata.Text = parent.gCurrency
			
			.vspddata.col = C_ExchRate
			.vspddata.Text = 1	

			.vspddata.col = C_deptcd
			.vspddata.Text = .txtDeptCd.value

			.vspddata.col = C_deptnm
			.vspddata.Text = .txtDeptNm.value

			
			Call MaxSpreadVal(.vspdData, C_ItemSeq, ii)
		Next

		.vspdData.Col = 1																	' 컬럼의 절대 위치로 이동      
		.vspdData.Row = ii - 1
		.vspdData.Action = 0
		Call CurFormatNumSprSheet()
        Call SetSpreadColorAcct(iCurRowPos + 1, iCurRowPos + imRow)
        .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   

'	ggoSpread.Source = .vspddata2
'	ggoSpread.ClearSpreadData		
		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related To DeleteRow ButTon of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
    Dim lDelRows 
    Dim DelItemSeq

	With frm1
		If gSelframeFlg = TAB1 Then
			If .vspddata4.Maxrows < 1 Then Exit Function
			ggoSpread.Source = .vspddata4
			
			lDelRows = ggoSpread.DeleteRow		
			Call DoSum()		
		Else
			If .vspddata.Maxrows < 1 Then Exit Function

			.vspddata.Row = .vspddata.ActiveRow
			.vspddata.Col = C_ItemSeq 
			DelItemSeq = .vspddata.Text

			ggoSpread.Source = .vspddata 
			lDelRows = ggoSpread.DeleteRow

			ggoSpread.Source = .vspdData2
			ggoSpread.ClearSpreadData		
			DeleteHsheet DelItemSeq
		End If
	End With

	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related To Print ButTon of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next  
    parent.FncPrint()      
    		
	Set gActiveElement = document.activeElement    
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)      
    		
	Set gActiveElement = document.activeElement    
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related To Excel 
'========================================================================================
Function  FncExcel() 
	Call FncExport(parent.C_SINGLEMULTI)
			
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

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2, var3
	
	FncExit = False

	ggoSpread.Source = frm1.vspddata4
    var1 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspddata
    var2 = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspddata2
    var3 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var1 = True or var2 = True or var3 = True Then  '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    FncExit = True
    		
	Set gActiveElement = document.activeElement    
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************




'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    DbDelete = False														

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtAllcNo=" & Trim(frm1.htxtAllcNo.value)				'☜: 삭제 조건 데이타 

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================================
Function DbDeleteOk()															'삭제 성공후 실행 로직 
	Call ggoOper.ClearField(Document, "1")								'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")								'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'Lock  Suitable  Field

	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspddata4
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData

    Call InitVariables()														'Initializes local global variables
    Call Clicktab1()    
    Call SetDefaultVal()
    Call DisableRefPop()

    frm1.txtAllcNo.Value = ""
    frm1.txtAllcNo.focus
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbQuery() 
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.htxtAllcNo.value)			'조회 조건 데이타 
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			'☜: 
			strVal = strVal & "&txtAllcNo=" & Trim(.txtAllcNo.value)			'조회 조건 데이타 
		End If
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)										    '☜: 비지니스 ASP 를 가동 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk()
	dim lngRows
	With frm1
	    Call SetSpreadLock("A")
	    Call SetSpreadLock("B")
        '-----------------------
        'Reset variables area
        '-----------------------       
        If .vspddata.MaxRows > 0 Then
            .vspddata.Row = 1
            .vspddata.Col = C_ItemSeq
            .hItemSeq.Value = .vspddata.Text 
            Call DbQuery2(1)
        End If
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field        
        Call InitVariables()
        lgIntFlgMode = parent.OPMD_UMODE										'Indicates that current mode is Update mode
        Call ClickTab1()
    End With 

	lgQueryOk = True
	
'	Call DoSum()	

'	Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_OPEN_AMT,"A", "I" ,"X","X")
'	Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_BAL_AMT,"A", "I" ,"X","X")
'	Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_CLS_AMT,"A", "I" ,"X","X")
'	Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_DC_AMT,"A", "I" ,"X","X")
'	Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_XCH_RATE,"D", "I" ,"X","X")			

'	Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DocCur,C_ItemAmt,"A", "I" ,"X","X")
'	Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DocCur,C_ExchRate,"D", "I" ,"X","X")

'    Call CurFormatNumSprSheet()

  

	With frm1.vspddata
	
		For lngRows=.MaxRows  To 1 step -1 
		
		.Row = lngRows
		.Col = 1
		.action = 0
			Call DbQuery2(lngRows)
		next
 

		
	end with


		
    Call DisableRefPop()

    lgQueryOk = False
	lgBlnFlgChgValue = False    
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim pAP010M 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal 
    Dim strDel

    DbSave = False 

    Call LayerShowHide(1)

    On Error Resume Next                                                   
	Err.Clear 

	frm1.txtFlgMode.value = lgIntFlgMode
	frm1.txtMode.value = parent.UID_M0002

    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data 연결 규칙 
    ' 0: Sheet명, 1: Flag , 2: Row위치, 3~N: 각 데이타 
    lGrpCnt = 1
    strVal = ""
    strDel = ""
    
    ggoSpread.Source = frm1.vspddata4

	With frm1.vspddata4
		For lngRows = 1 To .MaxRows
		    .Row = lngRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.DeleteFlag
'					strDel = strDel & "D" & parent.gColSep  					'☜: C=Create, Row위치 정보 
'			        .Col = C_OPEN_TYPE								'1
'			        strDel = strDel & Trim(.Text) & parent.gColSep
'			        .Col = C_OPEN_NO
'			        strDel = strDel & Trim(.Text) & parent.gColSep
'			        .Col = C_OPEN_GL_SEQ
'			        strDel = strDel & Trim(.Text) & parent.gColSep
'			        .Col = C_DOC_CUR
'			        strDel = strDel & Trim(.Text) & parent.gColSep
'			        .Col = C_CLS_AMT
'			        strDel = strDel & "" & parent.gColSep
'			        .Col = C_CLS_LOC_AMT
'			        strDel = strDel & "" & parent.gColSep
'			        .Col = C_DC_AMT
'			        strDel = strDel & "" & parent.gColSep
'			        .Col = C_DC_LOC_AMT
'			        strDel = strDel & "" & parent.gColSep
'			        .Col = C_ITEM_DESC
'			        strDel = strDel & "" & parent.gRowSep		                    
'			        lGrpCnt = lGrpCnt + 1
			    Case Else
					strVal = strVal & "C" & parent.gColSep  					'☜: C=Create, Row위치 정보 
			        .Col = C_OPEN_TYPE											'1
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_OPEN_NO
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_OPEN_GL_SEQ
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_DOC_CUR
			        strVal = strVal & Trim(.Text) & parent.gColSep			        
			        .Col = C_CLS_AMT
			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
			        .Col = C_CLS_LOC_AMT
			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
			        .Col = C_DC_AMT
			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
			        .Col = C_DC_LOC_AMT
			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
			        .Col = C_ITEM_DESC
			        strVal = strVal & Trim(.Text) & parent.gRowSep		                    
			        lGrpCnt = lGrpCnt + 1							        
'				Case ggoSpread.insertFlag
'					strVal = strVal & "C" & parent.gColSep  					'☜: C=Create, Row위치 정보 
'			        .Col = C_OPEN_TYPE											'1
'			        strVal = strVal & Trim(.Text) & parent.gColSep
'			        .Col = C_OPEN_NO
'			        strVal = strVal & Trim(.Text) & parent.gColSep
'			        .Col = C_OPEN_GL_SEQ
'			        strVal = strVal & Trim(.Text) & parent.gColSep
'			        .Col = C_DOC_CUR
'			        strVal = strVal & Trim(.Text) & parent.gColSep			        
'			        .Col = C_CLS_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_CLS_LOC_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_DC_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_DC_LOC_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_ITEM_DESC
'			        strVal = strVal & Trim(.Text) & parent.gRowSep		                    
'			        lGrpCnt = lGrpCnt + 1				
'				Case ggoSpread.UpdateFlag
'					strVal = strVal & "U" & parent.gColSep  					'☜: C=Create, Row위치 정보 
'			        .Col = C_OPEN_TYPE											'1
'			        strVal = strVal & Trim(.Text) & parent.gColSep
'			        .Col = C_OPEN_NO
'			        strVal = strVal & Trim(.Text) & parent.gColSep
'			        .Col = C_OPEN_GL_SEQ
'			        strVal = strVal & Trim(.Text) & parent.gColSep
'			        .Col = C_DOC_CUR
'			        strVal = strVal & Trim(.Text) & parent.gColSep			        			        
'			        .Col = C_CLS_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_CLS_LOC_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_DC_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_DC_LOC_AMT
'			        strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'			        .Col = C_ITEM_DESC
'			        strVal = strVal & Trim(.Text) & parent.gRowSep		                    
'			        lGrpCnt = lGrpCnt + 1
			End Select				
		Next
	End With	
	
	frm1.txtMaxRows.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread.value =  strVal									'Spread Sheet 내용을 저장 

    lGrpCnt = 1
    strVal = ""
    strDel = ""    

	ggoSpread.Source = frm1.vspddata

	With frm1.vspddata
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.DeleteFlag
'					strDel = strDel & "D" & parent.gColSep										'C=Create, Sheet가 2개 이므로 구별 
'					.Col = C_ItemSeq	'1
'					strDel = strDel & Trim(.Text) & parent.gColSep
'					strDel = strDel & "" & parent.gColSep					
'					.Col = C_deptcd	'1
'					strDel = strDel & "" & parent.gColSep					
'					.Col = C_AcctCd	'1
'					strDel = strDel & "" & parent.gColSep
'					.Col = C_DrCrFg	'1
'					strDel = strDel & "" & parent.gColSep
'					.Col = C_DocCur	'1
'					strDel = strDel & "" & parent.gColSep
'					.Col = C_ExchRate	'1
'					strDel = strDel & "" & parent.gColSep					
'					.Col = C_ItemAmt		'2
'					strDel = strDel & "" & parent.gColSep
'					.Col = C_ItemLocAmt		'3
'					strDel = strDel & "" & parent.gColSep
'					.Col = C_ItemDesc		'4
'					strDel = strDel & "" & parent.gRowSep						
'
'					lGrpCnt = lGrpCnt + 1
				Case Else					
					strVal = strVal & "C" & parent.gColSep											'C=Create, Sheet가 2개 이므로 구별 
					.Col = C_ItemSeq	'1
					strVal = strVal & Trim(.Text) & parent.gColSep
					strVal = strVal & Trim(frm1.hOrgChangeId.value) & parent.gColSep					
					.Col = C_deptcd	'1
					strVal = strVal & Trim(.Text) & parent.gColSep					
					.Col = C_AcctCd	'1
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_DrCrFg	'1
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_DocCur	'1
					strVal = strVal & Trim(.Text) & parent.gColSep			
					.Col = C_ExchRate	'1
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep										
					.Col = C_ItemAmt		'2
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					.Col = C_ItemLocAmt		'3
					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
					.Col = C_ItemDesc		'4
					strVal = strVal & Trim(.Text) & parent.gRowSep

					lGrpCnt = lGrpCnt + 1					
										
'				Case ggoSpread.InsertFlag
'					strVal = strVal & "C" & parent.gColSep											'C=Create, Sheet가 2개 이므로 구별 
'					.Col = C_ItemSeq	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep
'					strVal = strVal & Trim(frm1.hOrgChangeId.value) & parent.gColSep					
'					.Col = C_deptcd	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep					
'					.Col = C_AcctCd	'1
''					strVal = strVal & Trim(.Text) & parent.gColSep
'					.Col = C_DrCrFg	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep
'					.Col = C_DocCur	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep			
'					.Col = C_ExchRate	'1
'					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep										
''					.Col = C_ItemAmt		'2
'					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'					.Col = C_ItemLocAmt		'3
'					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'					.Col = C_ItemDesc		'4
'					strVal = strVal & Trim(.Text) & parent.gRowSep
'
'					lGrpCnt = lGrpCnt + 1
'				Case ggoSpread.UpdateFlag
'					strVal = strVal & "U" & parent.gColSep 											'C=Create, Sheet가 2개 이므로 구별 
'					.Col = C_ItemSeq	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep
'					strVal = strVal & Trim(frm1.hOrgChangeId.value) & parent.gColSep					
'					.Col = C_deptcd	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep					
'					.Col = C_AcctCd	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep
'					.Col = C_DrCrFg	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep
'					.Col = C_DocCur	'1
'					strVal = strVal & Trim(.Text) & parent.gColSep
'					.Col = C_ExchRate	'1
'					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep											
'					.Col = C_ItemAmt		'2
'					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'					.Col = C_ItemLocAmt		'3
'					strVal = strVal & Trim(UNIConvNum(.Text,0)) & parent.gColSep
'					.Col = C_ItemDesc		'4
'					strVal = strVal & Trim(.Text) & parent.gRowSep
'
'					lGrpCnt = lGrpCnt + 1					
			End Select							        
		Next
	End With
	
	frm1.txtMaxRows1.value = lGrpCnt-1										'Spread Sheet의 변경된 최대갯수 
	frm1.txtSpread1.value =  strVal									'Spread Sheet 내용을 저장    
				
    lGrpCnt = 1
    strVal = ""
    strDel = ""    

    With frm1.vspddata3   
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.DeleteFlag

				Case Else
					strVal = strVal & "C" & parent.gColSep 											'C=Create, Sheet가 2개 이므로 구별 
					.Col = 1	
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = 2
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = 3
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = 5
					strVal = strVal & Trim(.Text) & parent.gRowSep  
					lGrpCnt = lGrpCnt + 1		
			End Select
		Next
	End With

	With frm1
		.txtMaxRows3.value = lGrpCnt-1															'Spread Sheet의 변경된 최대갯수 
		.txtSpread3.value =  strVal																'Spread Sheet 내용을 저장 
		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'저장 비지니스 ASP 를 가동 

    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function  DbSaveOk(ByVal AllcNo)													'☆: 저장 성공후 실행 로직 
    ggoSpread.SSDeleteFlag 1
    
    If lgIntFlgMode = parent.OPMD_CMODE Then
		frm1.txtAllcNo.value = AllcNo
	End If	  
	
	Call ggoOper.ClearField(Document, "2")											'Clear Contents  Field
    frm1.txtAllcNo.focus
    Call InitVariables()															'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspddata4
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
	
	Call DBquery()
End Function


'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
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
			Call SetSpread2ColorCtrl(Row)
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	, " & .hItemSeq.Value & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_OPEN_DTL C (NOLOCK), A_OPEN_ITEM D (NOLOCK) "
		
		strWhere =			  " D.ITEM_NO =  " & FilterVar(UCase(.txtAllcNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.ITEM_SEQ = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.ITEM_NO  =  C.ITEM_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
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

			For lngRows = 1 To frm1.vspdData2.Maxrows
				frm1.vspddata2.Row = lngRows	
				frm1.vspddata2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspddata2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspddata2.Col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspddata2.Col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text, "''", "S") & "  " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If				 
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.Col = C_CtrlValNm
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
		
'		Call CopyFromData(.hItemSeq.value)
		Call SetSpread2ColorCtrl(Row)
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function  DbQueryOk2()
	Call SetSpread2ColorCtrl(1)
   
    lgBlnFlgChgValue = False        
End Function

'===================================== DisableRefPop()  =======================================
'	Name : DisableRefPop()
'	Description :
'====================================================================================================
Sub DisableRefPop()
	SpnRef.innerHTML="<A href=""vbscript:OpenRefOpenNo()""0>미결통합참조</A>"
End sub
'=======================================================================================================
' Function Name : chkAllcDate
' Function Desc : This function is related To Delete ButTon of Main ToolBar
'========================================================================================================
Function chkAllcDate()
	Dim intI
	
	chkAllcDate = True
	With frm1
		For intI = 1 To .vspddata4.Maxrows
			.vspddata4.Row = intI
			.vspddata4.Col = C_OPEN_DT		
			'반제일 
			If CompareDateByFormat(.vspddata4.Text,.txtAllcDt.Text,"미결발생일자",.txtAllcDt.Alt, _
		    	               "970025",.txtAllcDt.UserDefinedFormat,parent.gComDateType, True) = False Then
				.txtAllcDt.focus
				chkAllcDate = False
				Exit Function
			End If
		Next
	End With
End Function

'====================================================================================================
'	Name : DoSum()
'	Description : Sum Sheet Data
'====================================================================================================
Sub DoSum()
	Dim dblDrLocAmt
	Dim dblCrLocAmt

	Dim ii
	Dim iDocCur
	Dim iDrCrFg 
	
	dblDrLocAmt = 0
	dblCrLocAmt = 0
	
	With frm1
		If .vspddata4.MaxRows <> 0 Then
			For ii = 1 To .vspddata4.MaxRows
				.vspddata4.Row = ii		
				.vspddata4.Col = 0
				If .vspddata4.text <> ggoSpread.DeleteFlag Then
					.vspddata4.Col = C_DR_CR_FG
					iDrCrFg = UCase(Trim(.vspddata4.Text))
					.vspddata4.Col = C_DOC_CUR
					iDocCur = UCase(Trim(.vspddata4.Text))
					.vspddata4.Col = C_CLS_LOC_AMT
					If iDrCrFg = "DR" Then
						If .vspddata4.Text = "" Then			
							dblCrLocAmt = dblCrLocAmt +  0
						Else
							dblCrLocAmt = dblCrLocAmt +  UNICDBL(Trim(.vspddata4.Text))
						End If	
					Else
						If .vspddata4.Text = "" Then						
							dblDrLocAmt = dblDrLocAmt +  0
						Else
							dblDrLocAmt = dblDrLocAmt +  UNICDBL(Trim(.vspddata4.Text))
						End If	
					End If

					.vspddata4.Col = C_DC_LOC_AMT
					If iDrCrFg = "DR" Then
						If .vspddata4.Text = "" Then			
							dblCrLocAmt = dblCrLocAmt +  0
						Else
							dblCrLocAmt = dblCrLocAmt +  UNICDBL(Trim(.vspddata4.Text))
						End If	
					Else
						If .vspddata4.Text = "" Then						
							dblDrLocAmt = dblDrLocAmt +  0
						Else
							dblDrLocAmt = dblDrLocAmt +  UNICDBL(Trim(.vspddata4.Text))
						End If	
					End If
				End If				
			Next
		End If

		If .vspddata.MaxRows <> 0 Then
			For ii = 1 To .vspddata.MaxRows
			    .vspddata.Row = ii
			    .vspddata.Col = 0
			    If .vspddata.text <> ggoSpread.DeleteFlag Then		
					.vspddata.Col = C_DrCrFg
					iDrCrFg = UCase(Trim(.vspddata.Text))
					.vspddata.Col = C_DocCur
					iDocCur = UCase(Trim(.vspddata.Text))

					.vspddata.Col = C_ItemLocAmt				

					If iDrCrFg = "DR" Then
				        If .vspddata.Text = "" Then
							dblDrLocAmt = dblDrLocAmt +  0
						Else
							dblDrLocAmt = dblDrLocAmt +  UNICDBL(Trim(.vspddata.Text))
						End If	
					Else
				        If .vspddata.Text = "" Then
							dblCrLocAmt = dblCrLocAmt +  0
						Else				
							dblCrLocAmt = dblCrLocAmt +  UNICDBL(Trim(.vspddata.Text))
						End If	
					End If
				End If		
			Next		
		End If
		.txtDrLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblDrLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, "X", "X")
		.txtCrLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblCrLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, "X", "X")
		
		.txtDiffLocAmt.text = UNIConvNumPCToCompanyByCurrency(dblDrLocAmt-dblCrLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, "X", "X")
	End With
End Sub

'========================================================================================
' Function Name : FncBtnCalc
' Function Desc : This function calculate local amt from amt of multi
'========================================================================================
Function FncBtnCalc() 
	Dim ii
	Dim tempAmt, tempLocAmt, tempExch, TempSep, tempDoc , tempClsAmt, tempClsLocAmt,tempDcAmt, tempDcLocAmt
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strDate
	Dim strExchFg
	Dim IntRetCD
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	With frm1
		strSelect	= "b.minor_cd"
		strFrom		= "b_company a, b_minor b"
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  And	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If

		strDate = UniConvDateToYYYYMMDD(.txtAllcDt.text,parent.gDateFormat,"")

		If .vspdData4.MaxRows <> 0 Then
			For ii = 1 To .vspdData4.MaxRows
				.vspdData4.Row	=	ii
				.vspdData4.Col	=	C_DOC_CUR
				tempDoc			=	UCase(Trim(.vspdData4.text))
				.vspdData4.Col	=	C_CLS_AMT
				tempClsAmt		=	UNICDbl(.vspdData4.text)
				.vspdData4.Col	=	C_DC_AMT
				tempDcAmt		=	UNICDbl(.vspdData4.text)				
				.vspdData4.Col	=	C_XCH_RATE
				tempExch		=	UNICDbl(.vspdData4.text)

				If tempDoc	<> "" And tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox ("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "Top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep = arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox ("121500", "X", "X", "X")
						End If
					End If
 
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempClsLocAmt	=  tempClsAmt * TempExch
						tempDcLocAmt    =  tempDcAmt  * TempExch
					Else
						tempClsLocAmt	=  tempClsAmt / TempExch
						tempDcLocAmt	=  tempDcAmt / TempExch
					End If
					.vspdData4.Col	= C_CLS_LOC_AMT
					.vspdData4.text	= UNIConvNumPCToCompanyByCurrency(tempClsLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
					.vspdData4.Col	= C_DC_LOC_AMT
					.vspdData4.text	= UNIConvNumPCToCompanyByCurrency(tempDcLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")					
				ElseIf tempDoc = parent.gCurrency Then
					.vspdData4.Col	= C_CLS_LOC_AMT
					.vspdData4.text	= UNIConvNumPCToCompanyByCurrency(tempClsAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
					.vspdData4.Col	= C_DC_LOC_AMT
					.vspdData4.text	= UNIConvNumPCToCompanyByCurrency(tempDcAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")					
				End If
			Next
		End If



		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DocCur
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_ItemAmt
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_ExchRate
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" And tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox ("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "Top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And To_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep = arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox ("121500", "X", "X", "X")
						End If
					End If

					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt	=	tempAmt * TempExch
					Else
						tempLocAmt	=	tempAmt / TempExch
					End If
					.vspdData.Col	= C_ItemLocAmt
					.vspdData.text	= UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	= C_ItemLocAmt
					.vspdData.text	= UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmTofMoneyNo, parent.gLocRndPolicyNo, "X")
				End If
			Next
		End If
	End With

	Call DoSum
End Function

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_OPEN_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_BAL_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_CLS_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_DC_AMT,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData4,1,.vspddata4.MaxRows,C_DOC_CUR,C_XCH_RATE,"D", "I" ,"X","X")			

		Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DocCur,C_ItemAmt,"A", "I" ,"X","X")
		Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DocCur,C_ExchRate,"D", "I" ,"X","X")
	End With	
End Sub


'====================================================================================================
'	Name : XchLocRate()
'	Description : 환율이 변경되는 FacTor 가 변했을 때 수정되는 Local Amt. Setting
'====================================================================================================
Sub XchLocRate(ByVal SpdNo, ByVal Row)
	With frm1
		Select Case SpdNo
			Case "A"
		
			Case "B"
				.vspdData.Row = Row	
				.vspdData.Col = C_ItemLocAmt	
				.vspdData.Text = ""  
				ggoSpread.Source = .vspdData
				ggoSpread.UpdateRow Row  		
		End Select
	End With
End Sub

'*******************************************************************************************************
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

'===================================== PopResToreSpreadColumnInf()  ======================================
' Name : PopResToreSpreadColumnInf()
' Description : 컬럼의 순서정보를 복원함 
'====================================================================================================
Sub  PopResToreSpreadColumnInf()
	Dim indx

	ggoSpread.Source = gActiveSpdSheet
	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			Call PrevspdDataResTore(gActiveSpdSheet)
			Call ggoSpread.ResToreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpreadLock("B")
			Call SetSpread2ColorCtrl(1)									
		Case "VSPDDATA4" 
'			Call PrevspdDataResTore(gActiveSpdSheet)
			Call ggoSpread.ResToreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()			
			Call SetSpreadLock("A")
		Case "VSPDDATA2"
			Call PrevspdData2ResTore(gActiveSpdSheet)   
			Call ggoSpread.ResToreSpreadInf()
			Call InitCtrlSpread()			'관리항목 그리드 초기화 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2ColorCtrl(1)  
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If		
End Sub

'===================================== PrevspdDataResTore()  ========================================
' Name : PrevspdDataResTore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdDataResTore(pActiveSheetName)
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
									Call FncResToreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtAllcNo.Value)
								End If
						End Select
					Next
				Case ggoSpread.DeleteFlag
					Call fncResToreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtAllcNo.Value)
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'===================================== PrevspdDataResTore()  ========================================
' Name : PrevspdData2ResTore()
' Description : 그리드 복원시 관리항목 복원 
'====================================================================================================
Sub PrevspdData2ResTore(pActiveSheetName)
	Dim indx, indx1

	For indx = 0 To frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 To frm1.vspdData.MaxRows
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
					For indx1 = 0 To frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncResToreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtAllcNo.Value) 
						End If
					Next
				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	ggoSpread.Source = pActiveSheetName
End Sub

'========================================================================================================
' Name : fncResToreDbQuery2																				
' Desc : This function is data query and display												
'========================================================================================================
Function fncResToreDbQuery2(Row, CurrRow, Byval pInvalue1)
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

	fncResToreDbQuery2 = False

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
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	, " & strItemSeq & ",  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_RCPT_DC_DTL C (NOLOCK), A_RCPT_DC D (NOLOCK) "
		
		strWhere =			  " D.ALLC_NO =  " & FilterVar(UCase(.txtALLCNo.value), "''", "S") & "  "
		strWhere = strWhere & " AND D.SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.ALLC_NO  =  C.ALLC_NO  "
		strWhere = strWhere & " AND D.SEQ  =  C.SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
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
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call ResToreToolBar()
	End With

	If Err.number = 0 Then
		fncResToreDbQuery2 = True
	End If
End Function

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
'   Event Name : vspddata4_onfocus
'   Event Desc :
'=======================================================================================================
Sub  vspddata_onfocus()
    If lgIntFlgMode <> parent.OPMD_UMODE Then    
        Call SetToolBar("1110111100001111")                                     '버튼 툴바 제어 
    Else                 
        Call SetToolBar("1111111100001111")                                     '버튼 툴바 제어 
    End If  
End Sub

'==========================================================================================
'   Event Name : vspddata4_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspddata4_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"									'Split 상태코드 
 	Set gActiveSpdSheet = frm1.vspddata4
	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if

	If Row <= 0 Then
		Exit Sub
	End If		
End Sub

'==========================================================================================
'   Event Name : vspddata2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspddata_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("1101111111")
	
    gMouseClickStatus = "SP1C"									'Split 상태코드 
 	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 then
	    Exit Sub
	End if

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'AscEnding Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'DescEnding Sort
			lgSortKey = 1
		End If																
		Exit Sub
	End If		

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Row = frm1.vspdData.ActiveRow	
 	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) > 0 Then
	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData		
	End if	
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspddata4_MouseDown(ButTon, Shift, X, Y)
	If ButTon = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspddata_MouseDown(ButTon, Shift, X, Y)
	If ButTon = 2 And gMouseClickStatus = "SP1C" Then
		gMouseClickStatus = "SP1CR"
	End If
End Sub

'=======================================================================================================
'   Event Name : vspddata4_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspddata.Col = C_ItemSeq
            .hItemSeq.value = .vspddata.Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData		
        End With

        frm1.vspddata.Col = 0

        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
        End if

        lgCurrRow = NewRow

        Call DbQuery2(lgCurrRow)
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
'   Event Name : vspddata4_ColWidthChange
'   Event Desc : 상세내역 그리드의 (멀티)컬럼의 너비를 조절하는 경우 
'=======================================================================================================
Sub  vspddata4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspddata4
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'======================================================================================================
'   Event Name : vspddata4_ButTonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspddata_ButtonClicked(ByVal Col, ByVal Row, Byval ButTonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspddata
        ggoSpread.Source = frm1.vspddata
       
        If Row > 0 And (Col = C_AcctPopup  Or Col = C_deptPopup Or Col = C_DocCurPopup) Then
            .Col = Col - 1
            .Row = Row
			If Col = C_AcctPopup  Then
				Call OpenPopUp(.Text , 3)
			Elseif Col = C_deptPopup Then
				Call OpenUnderDept(.Text, 1)
			Else Col = C_DocCurPopup 
				Call OpenPopUp(.Text , 1)
			End If	
        End If    
    End With
End Sub



'======================================================================================================
'   Event Name :vspddata4_EditChange
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name : vspddata4_Change
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_Change(ByVal Col, ByVal Row )
	Dim OpenAmt
	Dim ClsAmt
	Dim iDocCur
	Dim DcAmt
	Dim dblTotClsAmt
	Dim dblTotDcAmt

	With frm1
		ggoSpread.Source = .vspddata4
		ggoSpread.UpdateRow Row
    
		.vspddata4.Row = Row
		.vspddata4.Col = Col
    
		Select Case Col
			Case C_CLS_AMT
				.vspddata4.Col = C_OPEN_AMT
				OpenAmt = .vspddata4.Text
				.vspddata4.Col = C_CLS_AMT
				ClsAmt =UNICDbl(.vspddata4.Text)

				.vspddata4.Col = C_CLS_LOC_AMT	
				.vspddata4.Text = ""

				If (UNICDbl(OpenAmt) > 0 And UNICDbl(ClsAmt) < 0) Or (UNICDbl(OpenAmt) < 0 And UNICDbl(ClsAmt) > 0) Then
					.vsppdata4.Col = C_DOC_CUR
					iDocCur = .vsppdata4.Text
				
					.vspddata4.Col = C_CLS_AMT
					.vspddata4.Text = UNIConvNumPCToCompanyByCurrency(ClsAmt * (-1),iDocCur,parent.ggAmTofMoneyNo, "X", "X")
				End If
				
				Call DoSum()				
			Case C_DC_AMT
				.vspddata4.Col = C_DC_LOC_AMT	
				.vspddata4.Text = ""			

				Call DoSum()
			Case C_CLS_LOC_AMT
				Call DoSum()
			Case C_DC_LOC_AMT				
				Call DoSum()				
		End Select
	End With	
End Sub




'==========================================================================================

Sub DeptCd_underChange(Byval stsfg, Byval strCode)
        

    Dim IntRetCD 
	dim strVal, vRow,strWhere
	
    lgBlnFlgChgValue = True
    
    strWhere=" AND A.COST_CD = B.COST_CD "
	strWhere= strWhere & " AND A.DEPT_CD="&filterVar(strCode,"''","S") & " AND A.END_DEPT_FG='Y' and  A.ORG_CHANGE_ID =  " & filterVar(parent.gChangeOrgId,"''","S") 
	IntRetCd =  CommonQueryRs(" DEPT_NM "," B_ACCT_DEPT A, B_COST_CENTER B "," 1=1 " & strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	vRow = frm1.vspdData.ActiveRow
	
	if IntRetCd=true then
		
			frm1.vspdData.Row = vRow
			frm1.vspdData.Col = C_deptcd + 2	
			frm1.vspdData.text =  Split(lgF0, chr(11))(0)
	else
			frm1.vspdData.Row = vRow
			frm1.vspdData.Col = C_deptcd + 2
			frm1.vspdData.text =  ""
	
	end if
	

End Sub




'======================================================================================================
'   Event Name : vspddata_Change
'   Event Desc :
'=======================================================================================================
Sub  vspddata_Change(ByVal Col, ByVal Row )
	Dim TempAmt

    With frm1
		ggoSpread.Source = .vspddata
		ggoSpread.UpdateRow Row

		.vspddata.Row = Row
		.vspddata.Col = 0
    
		Select Case Col
			Case   C_DeptCd
				.vspdData.Col = C_DeptCd
				Call DeptCd_underChange("",.vspdData.text)	
					
			Case  C_AcctCd
				If .vspddata.Text = ggoSpread.InsertFlag Then
				    .vspddata.Col = C_ItemSeq
				    .hItemSeq.value = .vspddata.Text
				    .vspddata.Col = C_AcctCd
				    If Len(.vspddata.Text) > 0 Then
						.vspddata.Row = Row
						.vspddata.Col = C_ItemSeq   	
						Call DeleteHsheet(.vspddata.Text)
				        Call DbQuery3(Row)
						Call SetSpread2ColorCtrl(Row)
				    End If    
				End If
				
	  		Case	C_DrCrFg
    			Call DoSum	
    		Case	C_DrCrNm
    			Call vspdData_ComboSelChange(Col,Row)
				Call DoSum	
    		Case   C_ItemAmt
				.vspdData.Row = Row
				.vspdData.Col = C_ItemAmt
				
    			TempAmt = UNICDbl(.vspdData.text)
    			    		
    			.vspdData.Row = Row
				.vspdData.Col = C_DocCur
    	
    			If UCase(Trim(.vspdData.Text)) = parent.gCurrency Then
					.vspdData.Row = Row
					.vspdData.Col = C_ItemLocAmt
					.vspdData.Text = TempAmt
				Else
					.vspdData.Row = Row
					.vspdData.Col = C_ItemLocAmt
					.vspdData.Text = ""
				End If
				
    			Call DoSum()
			Case   C_ItemLocAmt

				.vspdData.Row = Row
				.vspdData.Col = C_ItemAmt
				
    			TempAmt = UNICDbl(.vspdData.text)
    			    		
    			frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_DocCur
    	
    			If UCase(Trim(.vspdData.Text)) = parent.gCurrency Then
					.vspdData.Row = Row
					.vspdData.Col = C_ItemLocAmt
					.vspdData.Text = TempAmt
				End If

				Call DoSum()
			Case	C_ExchRate
				.vspdData.Row = Row
				.vspdData.Col = C_DocCur
				If UCase(Trim(.vspdData.Text)) = parent.gCurrency Then
					.vspdData.Row = Row
					.vspdData.Col = C_ExchRate
					.vspdData.Text = 1
				End If
				
				.vspdData.Col = C_ItemLocAmt
				.vspdData.Text = ""
			Case	C_DocCur
				.vspdData.Row  = Row
				.vspdData.Col  = C_ItemLocAmt
				.vspdData.Text = ""
				.vspdData.Col  = C_DocCur
				If UCase(Trim(.vspdData.Text)) = parent.gCurrency Then
					.vspdData.Col = C_ExchRate
					.vspdData.Text = 1
				Else
					Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtAllcDt.text,parent.gDateFormat,""), UCase(Trim(frm1.vspdData.Text)),frm1.vspdData.ActiveRow)
				End If
				
				Call DocCur_OnChange(Row,Row)
		End Select
	End With	
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim tmpDrCrFg
	Dim ii
	Dim iChkAcctForVat

	With frm1
		.vspddata.Row = Row
		Select Case Col
			Case C_DrCrNm
				.vspddata.Col = Col
				intIndex = .vspddata.Value
				.vspddata.Col = C_DrCrFg
				.vspddata.Value = intIndex
				tmpDrCrFg = .vspddata.text
				Call SetSpread2Color
			Case C_VatNm
				.vspddata.Col = Col
			    intIndex = .vspddata.Value
				.vspddata.Col = C_VatType
				.vspddata.Value = intIndex
			    Call InputCtrlVal(Row)'
		End Select
	End With
End Sub

'==========================================================================================
Sub DocCur_OnChange(ByVal FromRow, ByVal ToRow)
	Dim ii
	
    lgBlnFlgChgValue = True
	
	With frm1 
		For ii = FromRow	To	ToRow
			.vspdData.Row	= ii
			.vspdData.Col	= C_DocCur
			If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
				Call CurFormatNumSprSheet()
				Call DoSum
			End If
		Next
	End With
End Sub

'======================================================================================================
'   Event Name :vspddata4_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata4_DblClick
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_DblClick( ByVal Col , ByVal Row )
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspddata4.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("B")
End Sub

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspddata4 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name :vspddata4_KeyPress
'   Event Desc :
'=======================================================================================================
Sub  vspddata4_KeyPress(KeyAscii )
     
End Sub

'======================================================================================================
'   Event Name : vspddata4_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub  vspddata_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.7 Date-Numeric OCX Tag Event
' Description : This part declares HTML Tag Event
'=======================================================================================================
'*******************************************************************************************************

'=======================================================================================================
'   Event Name : txtAllcDt_DblClick(ButTon)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_DblClick(ButTon)
    If ButTon = 1 Then
        frm1.txtAllcDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtAllcDt.Focus     
    End If
End Sub

'=======================================================================================================
'   Event Name : txtAllcDt_Change()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub  txtAllcDt_Change()
    lgBlnFlgChgValue = True
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD	WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/tab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>추가계정등록</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/tab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right ><A HREF="VBSCRIPT:OpenPopupTempGL()">결의전표</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A>&nbsp;|&nbsp;<Span id="SpnRef"><a href="vbscript:OpenRefOpenNo()">미결통합참조</A></Span></TD>
					<TD	WIDTH=10>&nbsp;</TD>
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
								<TR>
									<TD CLASS="TD5" NOWRAP>반제번호</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAllcNo" ALT="반제번호" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU"><IMG align=Top name=btnCalType src="../../../CShared/image/btnPopup.gif"  TYPE="BUTToN" onclick="vbscript: Call OpenPopup(frm1.txtAllcNo.value,0)"></TD>
								</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=80>
					<TD WIDTH="100%">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>반제일</TD>
								<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtAllcDt" CLASS=FPDTYYYYMMDD tag="22" Title="FPDATETIME" ALT="반제일"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>거래처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=21NXXU" ALT="거래처"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBpCd(frm1.txtBpCd.value)"> <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="거래처명"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=22NXXU" ALT="부서"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.value, 0)"> <INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="24" ALT="부서명"></TD>
								<TD CLASS=TD5 NOWRAP>거래통화</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=22NXXU" ALT="거래통화"><IMG SRC="../../image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.value, 2)"></TD>								
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>결의전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="결의전표번호"> </TD>																						
								<TD CLASS="TD5" NOWRAP>회계전표번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=19 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="회계전표번호"></TD>								
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtDesc" SIZE=76 MAXLENGTH=256  tag=22NXXU ALT="비고"></TD>								
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData4 width="100%" TITLE="SPREAD" tag="2" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>											
							</TR>												
						</TABLE>		
					</DIV>
					<DIV ID="TabDiv" SCROLL=no>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" TITLE="SPREAD" tag="2" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" TITLE="SPREAD" tag="2" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>	
							</TR>
						</TABLE>		
					</DIV>
					</TD>
				</TR>
				<TR HEIGHT=20>
					<TD WIDTH="100%">
						<TABLE <%=LR_SPACE_TYPE_60%>>						
						<TD CLASS=TD5 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>자국금액계산</BUTTON>&nbsp;
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDrLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="차변자국금액" tag="24X2"></OBJECT>');</SCRIPT>&nbsp;/&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtCrLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="대변자국금액" tag="24X2"></OBJECT>');</SCRIPT></TD>						
						</TD>
						<TD CLASS=TD5 WIDTH=* align=right COLSPAN=2>차이금액(자국)&nbsp;
						<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDiffLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="차이금액(자국)" tag="24X2"></OBJECT>');</SCRIPT>
						</TD>						
						</TABLE>
					</TD>							
				</TR>
<!--				<TR HEIGHT=20>
					<TD WIDTH="100%">
						<TABLE <%=LR_SPACE_TYPE_60%>>		
							<TD CLASS=TD5 NOWRAP>자국통화(차변/대변)</TD>
							<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDrLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="자국발생(차변)" tag="24X2"></OBJECT>');</SCRIPT>&nbsp;/&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtCrLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="자국발생(대변)" tag="24X2"></OBJECT>');</SCRIPT></TD>
							<TD CLASS=TD5 NOWRAP>외국통화(차변/대변)</TD>
							<TD CLASS=TD6 NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtDrLocAmt2" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="외화발생(차변)" tag="24X2"></OBJECT>');</SCRIPT>&nbsp;/&nbsp;
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtCrLocAmt2" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="외화발생(대변)" tag="24X2"></OBJECT>');</SCRIPT></TD>
						</TABLE>
					</TD>					
				</TR>				-->
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>			
</TABLE>
<TEXTAREA Class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread1		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread2		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA Class=hidden name=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows1"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows2"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtAllcNo"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDocCur"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hDocCur2"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT TYPE=hidden CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspddata3 width="100%" tag="2" TABINDEX="-1"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

