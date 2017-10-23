<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111ma1
'*  4. Program Name         : 멀티컴퍼니수주등록 
'*  5. Program Desc         : 멀티컴퍼니수주등록-멀티 
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/24
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "ksm111mb1.asp"
Const BIZ_PGM_ID2 = "ksm111mb01.asp"
Const BIZ_PGM_SAVE_ID = "ksm111mb1.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL = "KSM111QA1"
											'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'상단 스프레드 
Dim C_CFMFLG 	         '선택 
Dim C_CFM_YN		 '확정여부 
Dim C_PO_NO              '발주번호(고객주문번호)
Dim C_IMPORT_FLG         '내외자구분(수입여부:Y외자(Import),N내자(Domestic))
Dim C_PO_CUR             '화폐(주문화폐)
Dim C_PO_DOC_AMT         '발주순금액(발주거래금액(NET))
Dim C_PO_VAT_DOC_AMT     '부가세금액(부가세거래금액(PO))
Dim C_PO_TOT_DOC_AMT	 '발주총금액(=발주순금액+부가세금액 즉,C_PO_DOC_AMT+C_PO_VAT_DOC_AMT)
Dim C_PO_VAT_TYPE_CD     '부가세유형 
Dim C_PO_VAT_TYPE_NM     '부가세유형명 
Dim C_PO_VAT_RT          '부가세율(PO)
Dim C_PO_PAY_METH_CD     '결제방법 
Dim C_PO_PAY_METH_NM     '결제방법명 
Dim C_PO_INCOTERMS_CD    '가격조건 
Dim C_PO_INCOTERMS_NM    '가격조건명 

'상단 히든 스프레드 
Dim C_PO_COMPANY    	'발주법인 
Dim C_SO_COMPANY       	'수주법인 



'하단 스프레드 
Dim C_PUMMOK_CD		'품목코드 
Dim C_PUMMOK_NM		'품목명 
Dim C_PUMMOK_GK		'품목규격 
Dim C_ITEM_CD           '고객품목코드 
Dim C_ITEM_NM           '고객품목명 
Dim C_ITEM_GK           '고객품목규격 
Dim C_PO_QTY            '발주수량 
Dim C_PO_UNIT		'발주단위 
Dim C_PO_PRC            '발주단가 
Dim C_PO_DOC_AMT2        '금액(발주거래금액(NET))
Dim C_DLVY_DT           '납기일 
Dim C_VAT_DOC_AMT       '부가세금액(부가세거래금액)
Dim C_PO_VAT_RATE       '부가세율 
Dim C_PO_VAT_TYPE_CD2    '부가세유형 
Dim C_PO_VAT_TYPE_NM2    '부가세유형명 

'하단 히든 스프레드 
Dim C_ParentPrNo 	'발주순번 
Dim C_ParentRowNo       '상위 row 번호 
Dim C_Flag              '자기 번호 



Dim lgSpdHdrClicked	'2003-03-01 Release 추가 
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim lgIntFlgModeM                 'Variable is for Operation Status

Dim lgStrPrevKeyM			'Multi에서 재쿼리를 위한 변수 
Dim lglngHiddenRows		'Multi에서 재쿼리를 위한 변수	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgSortKey1
Dim lgSortKey2

Dim IsOpenPop
Dim lsClickCfmYes
Dim lsClickCfmNo

Dim lgCurrRow
Dim strInspClass

Dim lgPageNo1
Dim EndDate, StartDate,CurrDate, iDBSYSDate
iDBSYSDate = "<%=GetSvrDate%>"
CurrDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate = UnIDateAdd("m", 1, CurrDate, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, CurrDate, parent.gDateFormat)

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
	lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
	lgIntGrpCount = 0						'initializes Group View Size

	lgStrPrevKey1 = ""						'initializes Previous Key
	lgStrPrevKey2 = ""						'initializes Previous Key

	lgLngCurRows = 0						'initializes Deleted Rows Count
	lgSortKey1 = 2
	lgSortKey2 = 2
	lgPageNo = 0
	lgPageNo1 = 0

End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtFrDt.Text = StartDate
	frm1.txtToDt.Text = CurrDate
	'수주일 
	frm1.txtSo_dt.Text = CurrDate
	'유효일 
	frm1.txtvalid_dt.Text = CurrDate

	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
	frm1.btnSjSelect.disabled = True
	frm1.btnSjDisSelect.disabled = True

	Call SetToolbar("1100000000001111")

    	frm1.txtSupplierCd.focus

	Set gActiveElement = document.activeElement
	Set gActiveSpdSheet = frm1.vspdData
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'=============================== 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		.ReDraw = false
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050126",,Parent.gAllowDragDropSpread

		.MaxCols = C_SO_COMPANY + 1
		.Col = .MaxCols:	.ColHidden = True
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck    C_CFMFLG, "선택",10,,,true
		ggoSpread.SSSetCheck    C_CFM_YN, "확정여부",10,,,true

		ggoSpread.SSSetEdit 	C_PO_NO,"발주번호"		,15	'발주번호(고객주문번호)
		ggoSpread.SSSetEdit 	C_IMPORT_FLG,"내외자구분"	,15    	'내외자구분(수입여부:Y외자(Import),N내자(Domestic))
		ggoSpread.SSSetEdit 	C_PO_CUR,"화폐"			,15    	'화폐(주문화폐)
		ggoSpread.SSSetFloat	C_PO_DOC_AMT,"발주순금액"	,		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"    	'발주순금액(발주거래금액(NET))
		ggoSpread.SSSetFloat	C_PO_VAT_DOC_AMT,"부가세금액"	,		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"    	'부가세금액(부가세거래금액(PO))
		ggoSpread.SSSetFloat	C_PO_TOT_DOC_AMT,"발주총금액"	,		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z" 	'발주총금액(=발주순금액+부가세금액 즉,C_PO_DOC_AMT+C_PO_VAT_DOC_AMT)
		ggoSpread.SSSetEdit 	C_PO_VAT_TYPE_CD,"부가세유형"	,20    	'부가세유형 
		ggoSpread.SSSetEdit 	C_PO_VAT_TYPE_NM,"부가세유형명"	,20    	'부가세유형명 
		ggoSpread.SSSetFloat	C_PO_VAT_RT,"부가세율"		,		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"    	'부가세율(PO)
		ggoSpread.SSSetEdit 	C_PO_PAY_METH_CD,"결제방법"	,20    	'결제방법 
		ggoSpread.SSSetEdit 	C_PO_PAY_METH_NM,"결제방법명"	,20    	'결제방법명 
		ggoSpread.SSSetEdit 	C_PO_INCOTERMS_CD,"가격조건"	,20    	'가격조건 
		ggoSpread.SSSetEdit 	C_PO_INCOTERMS_NM,"가격조건명"	,20    	'가격조건명 

		ggoSpread.SSSetEdit 	C_PO_COMPANY,"발주법인"	,20    	'발주법인 
		ggoSpread.SSSetEdit 	C_SO_COMPANY,"수주법인"	,20    	'수주법인 


		Call ggoSpread.SSSetColHidden(C_PO_COMPANY,	C_PO_COMPANY,	True)
		Call ggoSpread.SSSetColHidden(C_SO_COMPANY,	C_SO_COMPANY,	True)

		Call SetSpreadLock

	    	.ReDraw = true
    	End With


End Sub

Sub InitSpreadSheet2()
	Call InitSpreadPosVariables2()



	With frm1.vspdData2
		.ReDraw = false
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20050126",,Parent.gAllowDragDropSpread

		.MaxCols = C_Flag+1
		.Col = .MaxCols:	.ColHidden = True

		.MaxRows = 0

		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit 	C_PUMMOK_CD	,"품목"  		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_PUMMOK_NM	,"품목명"  		,		25,		,					,	  40
		ggoSpread.SSSetEdit 	C_PUMMOK_GK	,"품목규격"  	,			20
		ggoSpread.SSSetEdit 	C_ITEM_CD       ,"고객품목"  	,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_ITEM_NM       ,"고객품목명"  	,		25,		,					,	  40
		ggoSpread.SSSetEdit 	C_ITEM_GK       ,"고객품목규격" 	,			20
		ggoSpread.SSSetFloat 	C_PO_QTY        ,"수량"  		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_PO_UNIT	,"단위"  		,			8,		,					,	  3,	  2
		ggoSpread.SSSetFloat 	C_PO_PRC        ,"단가"  		,			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_DOC_AMT2    ,"금액"  		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetDate 	C_DLVY_DT       ,"납기일"  		,		10,		2,					parent.gDateFormat
		ggoSpread.SSSetFloat 	C_VAT_DOC_AMT   ,"부가세금액"  	,		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_VAT_RATE   ,"부가세율"  	,		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_PO_VAT_TYPE_CD2,"부가세유형"  	,		10,		,					,	  5,	  2
		ggoSpread.SSSetEdit 	C_PO_VAT_TYPE_NM2,"부가세유형명" 	,	20

		ggoSpread.SSSetEdit 	C_ParentPrNo, "요청번호", 10
		ggoSpread.SSSetEdit	    C_ParentRowNo , "C_ParentRowNo", 5
		ggoSpread.SSSetEdit	    C_Flag , "C_Flag", 5


		Call ggoSpread.SSSetColHidden(C_ParentPrNo, C_ParentPrNo,	True)
		Call ggoSpread.SSSetColHidden(C_ParentRowNo,C_ParentRowNo, True)
 		Call ggoSpread.SSSetColHidden(C_Flag, C_Flag+1, True)

		Call SetSpreadLock2()
		.ReDraw = True

    	End With
End Sub

'============================= 2.2.4 SetSpreadLock() ====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
 Sub SetSpreadLock()
	With frm1.vspdData
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData

		ggoSpread.SpreadLock		C_PO_NO                 , -1, -1
		ggoSpread.SpreadLock		C_IMPORT_FLG            , -1, -1
		ggoSpread.SpreadLock		C_PO_CUR                , -1, -1
		ggoSpread.SpreadLock		C_PO_DOC_AMT            , -1, -1
		ggoSpread.SpreadLock		C_PO_VAT_DOC_AMT        , -1, -1
		ggoSpread.SpreadLock		C_PO_TOT_DOC_AMT	, -1, -1
		ggoSpread.SpreadLock		C_PO_VAT_TYPE_CD        , -1, -1
		ggoSpread.SpreadLock		C_PO_VAT_TYPE_NM        , -1, -1
		ggoSpread.SpreadLock		C_PO_VAT_RT             , -1, -1
		ggoSpread.SpreadLock		C_PO_PAY_METH_CD        , -1, -1
		ggoSpread.SpreadLock		C_PO_PAY_METH_NM        , -1, -1
		ggoSpread.SpreadLock		C_PO_INCOTERMS_CD       , -1, -1
		ggoSpread.SpreadLock		C_PO_INCOTERMS_NM       , -1, -1
		.ReDraw = True
	End With
End Sub

Sub SetSpreadLock2()
	With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2

		ggoSpread.SpreadLock		C_PUMMOK_CD		, -1, -1
                ggoSpread.SpreadLock		C_PUMMOK_NM	        , -1, -1
                ggoSpread.SpreadLock		C_PUMMOK_GK	        , -1, -1
                ggoSpread.SpreadLock		C_ITEM_CD               , -1, -1
                ggoSpread.SpreadLock		C_ITEM_NM               , -1, -1
                ggoSpread.SpreadLock		C_ITEM_GK               , -1, -1
                ggoSpread.SpreadLock		C_PO_QTY                , -1, -1
                ggoSpread.SpreadLock		C_PO_UNIT	        , -1, -1
                ggoSpread.SpreadLock		C_PO_PRC                , -1, -1
                ggoSpread.SpreadLock		C_PO_DOC_AMT2            , -1, -1
                ggoSpread.SpreadLock		C_DLVY_DT               , -1, -1
                ggoSpread.SpreadLock		C_VAT_DOC_AMT           , -1, -1
                ggoSpread.SpreadLock		C_PO_VAT_RATE           , -1, -1
                ggoSpread.SpreadLock		C_PO_VAT_TYPE_CD2        , -1, -1
                ggoSpread.SpreadLock		C_PO_VAT_TYPE_NM2        , -1, -1

		.ReDraw = True
    	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	With frm1.vspdData
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData

		ggoSpread.SSSetProtected	C_PO_NO                 , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_IMPORT_FLG            , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_CUR                , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_DOC_AMT            , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_DOC_AMT        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_TOT_DOC_AMT	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_TYPE_CD        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_TYPE_NM        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_RT             , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_PAY_METH_CD        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_PAY_METH_NM        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_INCOTERMS_CD       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_INCOTERMS_NM       , pvStartRow, pvEndRow

		.ReDraw = True
	End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
	With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2

		ggoSpread.SSSetProtected	C_PUMMOK_CD	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PUMMOK_NM	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PUMMOK_GK	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ITEM_CD       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ITEM_NM       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_ITEM_GK       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_QTY        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_UNIT	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_PRC        , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_DOC_AMT2    , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_DLVY_DT       , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_VAT_DOC_AMT   , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_RATE   , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_TYPE_CD2, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected	C_PO_VAT_TYPE_NM2, pvStartRow, pvEndRow

		.ReDraw = True
	End With
End Sub

'============================= 2.2.3 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_CFMFLG 	   	= 1     '선택 
	C_CFM_YN		= 2  	'확정여부 
	C_PO_NO             	= 3 	'발주번호(고객주문번호)
	C_IMPORT_FLG        	= 4 	'내외자구분(수입여부:Y외자(Import),N내자(Domestic))
	C_PO_CUR            	= 5 	'화폐(주문화폐)
	C_PO_DOC_AMT        	= 6 	'발주순금액(발주거래금액(NET))
	C_PO_VAT_DOC_AMT    	= 7 	'부가세금액(부가세거래금액(PO))
	C_PO_TOT_DOC_AMT	= 8 	'발주총금액(=발주순금액+부가세금액 즉,C_PO_DOC_AMT+C_PO_VAT_DOC_AMT)
	C_PO_VAT_TYPE_CD    	= 9 	'부가세유형 
	C_PO_VAT_TYPE_NM    	= 10 	'부가세유형명 
	C_PO_VAT_RT         	= 11 	'부가세율(PO)
	C_PO_PAY_METH_CD    	= 12 	'결제방법 
	C_PO_PAY_METH_NM    	= 13 	'결제방법명 
	C_PO_INCOTERMS_CD   	= 14 	'가격조건 
	C_PO_INCOTERMS_NM   	= 15 	'가격조건명 

	'상단 히든 스프레드 
	C_PO_COMPANY		=16    	'발주법인 
	C_SO_COMPANY		=17     '수주법인 


End Sub

Sub InitSpreadPosVariables2()
	C_PUMMOK_CD	    	= 1	'품목코드 
	C_PUMMOK_NM	    	= 2	'품목명 
	C_PUMMOK_GK	    	= 3	'품목규격 
	C_ITEM_CD           	= 4    	'고객품목코드 
	C_ITEM_NM           	= 5    	'고객품목명 
	C_ITEM_GK		= 6
	C_PO_QTY            	= 7    	'발주수량 
	C_PO_UNIT	    	= 8	'발주단위 
	C_PO_PRC            	= 9    	'발주단가 
	C_PO_DOC_AMT2        	= 10    	'금액(발주거래금액(NET))
	C_DLVY_DT           	= 11    '납기일 
	C_VAT_DOC_AMT       	= 12    '부가세금액(부가세거래금액)
	C_PO_VAT_RATE       	= 13    '부가세율 
	C_PO_VAT_TYPE_CD2    	= 14    '부가세유형 
	C_PO_VAT_TYPE_NM2    	= 15    '부가세유형명 

	'하단 히든 스프레드 
	C_ParentPrNo 		= 16'발주순번 
	C_ParentRowNo           = 17 '상위 row 번호 
	C_Flag                  = 18 '자기 번호 


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

			C_CFMFLG 	   	= iCurColumnPos(1)    	'선택 
			C_CFM_YN		= iCurColumnPos(2) 	'확정여부 
			C_PO_NO             	= iCurColumnPos(3)	'발주번호(고객주문번호)
			C_IMPORT_FLG        	= iCurColumnPos(4)	'내외자구분(수입여부:Y외자(Import),N내자(Domestic))
			C_PO_CUR            	= iCurColumnPos(5)	'화폐(주문화폐)
			C_PO_DOC_AMT        	= iCurColumnPos(6)	'발주순금액(발주거래금액(NET))
			C_PO_VAT_DOC_AMT    	= iCurColumnPos(7)	'부가세금액(부가세거래금액(PO))
			C_PO_TOT_DOC_AMT	= iCurColumnPos(8)	'발주총금액(=발주순금액+부가세금액 즉,C_PO_DOC_AMT+C_PO_VAT_DOC_AMT)
			C_PO_VAT_TYPE_CD    	= iCurColumnPos(9)	'부가세유형 
			C_PO_VAT_TYPE_NM    	= iCurColumnPos(10) 	'부가세유형명 
			C_PO_VAT_RT         	= iCurColumnPos(11) 	'부가세율(PO)
			C_PO_PAY_METH_CD    	= iCurColumnPos(12) 	'결제방법 
			C_PO_PAY_METH_NM    	= iCurColumnPos(13) 	'결제방법명 
			C_PO_INCOTERMS_CD   	= iCurColumnPos(14) 	'가격조건 
			C_PO_INCOTERMS_NM   	= iCurColumnPos(15) 	'가격조건명 

			C_PO_COMPANY   		= iCurColumnPos(16) 	''발주법인 
			C_SO_COMPANY   		= iCurColumnPos(17) 	'수주법인 


		Case "B"
			ggoSpread.Source = frm1.vspdData2
            		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PUMMOK_CD	    	= iCurColumnPos(1)	'품목코드 
			C_PUMMOK_NM	    	= iCurColumnPos(2)	'품목명 
			C_PUMMOK_GK	    	= iCurColumnPos(3)	'품목규격 
			C_ITEM_CD           	= iCurColumnPos(4)   	'고객품목코드 
			C_ITEM_NM           	= iCurColumnPos(5)   	'고객품목명 
			C_ITEM_GK		= iCurColumnPos(6)   	'고객품목규격 
			C_PO_QTY            	= iCurColumnPos(7)   	'발주수량 
			C_PO_UNIT	    	= iCurColumnPos(8)	'발주단위 
			C_PO_PRC            	= iCurColumnPos(9)   	'발주단가 
			C_PO_DOC_AMT2        	= iCurColumnPos(10)   	'금액(발주거래금액(NET))
			C_DLVY_DT           	= iCurColumnPos(11)    	'납기일 
			C_VAT_DOC_AMT       	= iCurColumnPos(12)    	'부가세금액(부가세거래금액)
			C_PO_VAT_RATE       	= iCurColumnPos(13)    	'부가세율 
			C_PO_VAT_TYPE_CD2    	= iCurColumnPos(14)    	'부가세유형 
			C_PO_VAT_TYPE_NM2    	= iCurColumnPos(15)    	'부가세유형명 


			C_ParentPrNo   		= iCurColumnPos(16) 	'상위 요청번호 (키값)
			C_ParentRowNo   	= iCurColumnPos(17) 	'상위 row 번호 
			C_Flag   		= iCurColumnPos(18) 	'자기 번호 

	End Select
End Sub

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 그리드의 숫자 부분이 변경된면 이 함수를 변경 해야함.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select
End Sub




'=======================================================================================================
' Function Name : DefaultCheck
' Function Desc :
'=======================================================================================================
Function DefaultCheck()
	DefaultCheck = False
	Dim i
	Dim j
	Dim RequiredColor

	ggoSpread.Source = frm1.vspdData2
	RequiredColor = ggoSpread.RequiredColor
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				.Col = 0
				If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Then
					For j = 1 To .MaxCols
						.Col = j
						If .BackColor = RequiredColor Then
							If Len(Trim(.Text)) < 1 Then
								.Row = 0
								Call DisplayMsgBox("970021","X",.Text,"")
								.Row = i
								.Action = 0
								Exit Function
							End If
						End If
					Next
				End If
			End If
		Next
	End With
	DefaultCheck = True
End Function

'=======================================================================================================
' Function Name : ChangeCheck
' Function Desc :
'=======================================================================================================
Function ChangeCheck()
	ChangeCheck = False

	Dim i
	Dim strInsertMark
	Dim strDeleteMark
	Dim strUpdateMark

	ggoSpread.Source = frm1.vspdData2
	strInsertMark = ggoSpread.InsertFlag
	strDeleteMark = ggoSpread.UpdateFlag
	strUpdateMark = ggoSpread.DeleteFlag

	If frm1.vspdData.maxrows <= 0 Then Exit Function
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = strInsertMark Or .Text = strDeleteMark Or .Text = strUpdateMark Then
				ChangeCheck = True
				exit for
			End If
		Next
	End With

	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or ChangeCheck = True Then
        ChangeCheck = True
    End If
End Function

'=======================================================================================================
' Function Name : CheckDataExist
' Function Desc :
'=======================================================================================================
Function CheckDataExist()
	CheckDataExist = False
	Dim i

	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataFirstRow
' Function Desc :
'=======================================================================================================
Function ShowDataFirstRow()
	ShowDataFirstRow = 0
	Dim i

	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : ShowDataLastRow
' Function Desc :
'=======================================================================================================
Function ShowDataLastRow()
	ShowDataLastRow = 0
	Dim i

	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function



'=======================================================================================================
' Function Name : DataFirstRow
' Function Desc :
'=======================================================================================================
Function DataFirstRow(ByVal Row)
	DataFirstRow = 0
	Dim i
	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataFirstRow = i
				Exit Function
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : DataLastRow
' Function Desc :
'=======================================================================================================
Function DataLastRow(ByVal Row)
	DataLastRow = 0
	Dim i

	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			.Col = C_ParentRowNo
			If Clng(.text) = Clng(Row) Then
				DataLastRow = i
				Exit Function
			End If
		Next
	End With
End Function

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc :
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If y<20 Then			'2003-03-01 Release 추가 
	    lgSpdHdrClicked = 1
	End If

    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================
' Function Name : vspdData2_MouseDown
' Function Desc :
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
    End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###

 	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0101111111")         '화면별 설정 
	End If

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		lgSpdHdrClicked = 0		'2003-03-01 Release 추가 
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)

	 	'------ Developer Coding part (End)
 	End If


End Sub

'================ vspdData_LeaveCell() ==========================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData

    If Row >= NewRow Then
        Exit Sub
    End If

    If NewRow = .MaxRows Then
        'DbQuery
    End if

    End With

End Sub



'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
 	Dim strShowDataFirstRow
 	Dim strShowDataLastRow
 	Dim i,k
 	Dim strFlag,strFlag1
 	gMouseClickStatus = "SP2C"

 	Set gActiveSpdSheet = frm1.vspdData2

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("1101111111")         '화면별 설정 
	End If

 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2
 		strShowDataFirstRow = Clng(ShowDataFirstRow)
 		strShowDataLastRow = Clng(ShowDataLastRow)
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col, lgSortKey2, strShowDataFirstRow, strShowDataLastRow	'Sort in Ascending
 			lgSortKey2 = 2
 		ElseIf lgSortKey2 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey2, strShowDataFirstRow, strShowDataLastRow	'Sort in Descending
 			lgSortKey2 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'========================================================================================
' Function Name : vspdData2_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD


	If Kubun = 0 Then

'		strTemp = ReadCookie("PoNo")

'		If strTemp = "" then Exit Function

'		frm1.txtPoNo.value = strTemp

'		WriteCookie "PoNo" , ""

'		Call MainQuery()

	elseIf Kubun = 1 Then

'	    If lgIntFlgMode <> Parent.OPMD_UMODE Then
'	        Call DisplayMsgBox("900002", "X", "X", "X")
'	        Exit Function
'	    End If

	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		With frm1
			'수주법인 
			WriteCookie "txtSoCompanyCd", Trim(.txtSupplierCd.value)
			WriteCookie "txtFrDt", Trim(.txtFrDt.text)
			WriteCookie "txtToDt", Trim(.txtToDt.text)
			WriteCookie "txtPO_NO", Trim(.txtPO_NO.value)

		End With

		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)
	End IF
End Function
'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()	'###그리드 컨버전 주의부분###

	Call LoadInfTB19029                                                         'Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	'Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")  'Lock  Suitable  Field
	Call InitSpreadSheet
	Call InitSpreadSheet2

	Call InitVariables
	Call SetDefaultVal
        Call CookiePage(0)

	Set gActiveSpdSheet = frm1.vspdData
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)

End Sub

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 그리드 열고정을 한다.
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
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
	Dim iActiveRow
	Dim iConvActiveRow
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim lRow
	Dim i
	Dim strFlag
	Dim strParentRowNo

    ggoSpread.Source = gActiveSpdSheet
    If gActiveSpdSheet.Name = "vspdData" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet
		Call ggoSpread.ReOrderingSpreadData

    ElseIf gActiveSpdSheet.Name = "vspdData2" Then
		For i = 1 To frm1.vspdData2.MaxRows
			frm1.vspdData2.Row = i
			frm1.vspdData2.Col = 0
			strFlag = frm1.vspdData2.Text
			If strFlag = ggoSpread.InsertFlag Then
				frm1.vspdData2.Col = C_ParentRowNo
				strParentRowNo = CInt(frm1.vspdData2.Text)
				lglngHiddenRows(strParentRowNo - 1) = CInt(lglngHiddenRows(strParentRowNo - 1)) - 1
			End If
		Next

		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet2
		frm1.vspdData2.Redraw = False

		Call ggoSpread.ReOrderingSpreadData("F")

		Call DbQuery2(frm1.vspdData.ActiveRow,False)

		lngRangeFrom = Clng(ShowDataFirstRow)
		lngRangeTo = Clng(ShowDataLastRow)

		lRow = frm1.vspdData.ActiveRow	'###그리드 컨버전 주의부분###
		frm1.vspdData2.Redraw = True
    End If

 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	'###그리드 컨버전 주의부분###
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release 추가 
		Exit Sub
	End If

	'/* 9월 정기패치 : 동일한 키값을 입력한 채 다른 스프레드로 옮기지 못하도록 수정관련 변수 추가 - START */
	Dim lRow
	'/* 9월 정기패치 : 동일한 키값을 입력한 채 다른 스프레드로 옮기지 못하도록 수정관련 변수 추가 - END */

	Set gActiveSpdSheet = frm1.vspdData

	frm1.vspdData.redraw = false
	If Row <> NewRow And NewRow > 0 Then
		With frm1
			.vspdData.redraw = false
			'/* 8월 정기패치 : 우측 스프레드에 필수입력 필드 체크 - START */
		'	If DefaultCheck = False Then
		'		.vspdData.Row = Row
		'		.vspdData.Col = 1
		'		.vspdData2.focus
    	'		Exit Sub
		'	End If
			'/* 8월 정기패치 : 우측 스프레드에 필수입력 필드 체크 - END */

			'/* 9월 정기패치: '다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - START */
			If CheckRunningBizProcess = True Then
				.vspdData.Row = Row
				.vspdData.Col = 1
				Exit Sub
			End If
			'/* 9월 정기패치: '다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - END */
			lgCurrRow = NewRow
			.vspdData.redraw = true
		End With

		lgIntFlgModeM = Parent.OPMD_CMODE

		With frm1.vspdData2
			.ReDraw = False
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.RowHidden = True
			.BlockMode = False
			.ReDraw = True
		End With
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData

		If DbQuery2(lgCurrRow, False) = False Then	Exit Sub
	End If
	frm1.vspdData.redraw = true
End Sub

'=======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	'/요청량이 변경되면 배부량을 수정한다.(요청량 * 배부비율)
	.Row = Row


    End With

End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt
    Dim LngLastRow
    Dim LngMaxRow

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    '/* 9월 정기패치: 해상도에 상관없이 재쿼리되도록 수정 - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '☜: 재쿼리 체크 
    '/* 9월 정기패치: 해상도에 상관없이 재쿼리되도록 수정 - END */
		If lgPageNo <> "" Then			'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If


			If DbQuery = False Then
				Exit Sub
			End If
		End If

    End If
End Sub

'======================================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt
    Dim LngLastRow
    Dim LngMaxRow
    Dim lRow

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    With frm1

    	lRow = .vspdData.ActiveRow
    	'/* 9월 정기패치: 해상도에 상관없이 재쿼리되도록 수정 - START */
    	If ShowDataLastRow < NewTop + VisibleRowCnt(.vspdData2, NewTop) Then	        '☜: 재쿼리 체크 
		'/* 9월 정기패치: 해상도에 상관없이 재쿼리되도록 수정 - END */
'    		If lgStrPrevKeyM(lRow - 1) <> "" Then            '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
    		If lgPageNo1 <> "" Then            '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				If CheckRunningBizProcess = True Then
					Exit Sub
				End If

				Call DisableToolBar(Parent.TBC_QUERY)
				If DbQuery2(lRow, True) = False Then
					Call RestoreToolBar()
					Exit Sub
				End If
			End If
		End If
    End With
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim index
    Dim intSeq

	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
	If Col = C_CFMFLG And Row > 0 Then
		frm1.vspdData.Redraw = false

		.Col = C_CFMFLG
		.Row = Row
		if Row <= 0 Then Exit Sub
	    If Trim(.value)="1" Then
			ggoSpread.UpdateRow Row
	    Else

		.Col  = 0
		.Row  = Row
		.text = ""
	    End If
		frm1.vspdData.Redraw = true
    	End If

	If Col = C_CFM_YN And Row > 0 Then
		frm1.vspdData.Redraw = false

		frm1.vspdData.Col = C_CFMFLG
		frm1.vspdData.Row = .ActiveRow
		if frm1.vspdData.value = 0 then
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_CFMFLG, .ActiveRow, 1)
		end if



		frm1.vspdData.Redraw = true
    End If




	End With

End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData2

		ggoSpread.Source = frm1.vspdData2

		If Row > 0 And Col = C_SpplPopup Then
			Call OpenSSupplier()
		Elseif Row > 0 And Col = C_GrpPopup Then
			Call OpenSGrp()
		End If

	End With
End Sub

'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================
'==========================================================================================
'   Event Name : txtFrDt
'   Event Desc : 고객발주일 
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtToDt
'   Event Desc : 고객발주일 
'==========================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtToDt.Focus
	End If
End Sub


'==========================================================================================
'   Event Name : txtSo_dt
'   Event Desc : 수주일 
'==========================================================================================
 Sub txtSo_dt_DblClick(Button)
	if Button = 1 then
		frm1.txtSo_dt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtSo_dt.Focus
	End If
End Sub


'==========================================================================================
'   Event Name : txtvalid_dt
'   Event Desc : 유효일 
'==========================================================================================
 Sub txtvalid_dt_DblClick(Button)
	if Button = 1 then
		frm1.txtvalid_dt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtvalid_dt.Focus
	End If
End Sub


'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================

Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


'======================================================================================================
' Function Name :
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###그리드 컨버전 주의부분###
	FncQuery = False
	Dim IntRetCD
	'-----------------------
	'Check previous data area
	'-----------------------
	If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
	'Erase contents area
	'-----------------------
	Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
	ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then											'This function check indispensable field
		Exit Function
	End If



 	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) and trim(.txtFrDt.text)<>""  then
			Call DisplayMsgBox("17a003", "X","고객발주일", "X")
			Exit Function
		End If

	End with

	'-----------------------
   	'Query function call area
    	'-----------------------

	If DbQuery = False then
		Exit Function
	End If

	Set gActiveElement = document.activeElement
    	FncQuery = True
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave()
    FncSave = False

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("181216", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If

<%  '-----------------------
    'Check content area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then    'Not chkField(Document, "2") OR      '⊙: Check contents area
       Exit Function
    End If


<%  '-----------------------
    'Check content area
    '----------------------- %>
	If Trim(UNIConvDateToYYYYMMDD(frm1.txtSo_dt.Text, gDateFormat,"")) > Trim(UNIConvDateToYYYYMMDD(frm1.txtvalid_dt.Text, gDateFormat,"")) Then
		Call DisplayMsgBox("970023","X", "유효일","수주일")
		Exit Function
	End If
With frm1
	'수주형태 
	if Trim(.txtSo_Type.value)=""  then
		Call DisplayMsgBox("17a003", "X","수주형태", "X")
		Exit Function
	End If


	'수주일		txtSo_dt
	if Trim(.txtSo_dt.value)=""  then
		Call DisplayMsgBox("17a003", "X","수주일", "X")
		Exit Function
	End If
	'판매유형	txtDeal_Type
	if Trim(.txtDeal_Type.value)=""  then
		Call DisplayMsgBox("17a003", "X","판매유형", "X")
		Exit Function
	End If
	'영업그룹	txtSales_Grp
	if Trim(.txtSales_Grp.value)=""  then
		Call DisplayMsgBox("17a003", "X","영업그룹", "X")
		Exit Function
	End If
	'공장	txtPlantCd
	if Trim(.txtPlantCd.value)=""  then
		Call DisplayMsgBox("17a003", "X","공장", "X")
		Exit Function
	End If
	'유효일 txtvalid_dt
	if Trim(.txtvalid_dt.value)=""  then
		Call DisplayMsgBox("17a003", "X","유효일", "X")
		Exit Function
	End If

End with



<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data

    FncSave = True
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
 	FncExcel = True
 End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
	FncPrint = False
	Call Parent.FncPrint()
	Set gActiveElement = document.activeElement
	FncPrint = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
	FncFind = False
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
	Set gActiveElement = document.activeElement
    FncFind = True
End Function


'=======================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()
	FncExit = False

	Dim IntRetCD

    If ChangeCheck = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery()
	DbQuery = False
	Dim strVal

	Call LayerShowHide(1)

	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then

			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '발주법인코드 
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)			'고객발주일 From
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)			'고객발주일 To

			if .rdoPostFlag2(0).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "N"
			elseif .rdoPostFlag2(1).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			else
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			End if

			strVal = strVal & "&txtPO_NO=" & Trim(.txtPO_NO.value)
			strVal = strVal & "&lgPageNo=" & lgPageNo                  '☜: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	Else

	        	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '발주법인코드 
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)			'고객발주일 From
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)			'고객발주일 To

			if .rdoPostFlag2(0).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "N"
			elseif .rdoPostFlag2(1).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			else
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			End if

			strVal = strVal & "&txtPO_NO=" & Trim(.txtPO_NO.value)
			strVal = strVal & "&lgPageNo=" & lgPageNo                  '☜: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

	    	End If
	End with

	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 

	DbQuery = True
End Function

'========================================================================================================
Function FncCancel()
	Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData

    Call CancelSum()

    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------

    '------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)
	DbQueryOk = False

	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows

	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("11001001000011")				'버튼 툴바 제어 

	frm1.btnSelect.disabled = false
	frm1.btnDisSelect.disabled = false
	frm1.btnSjSelect.disabled = false
	frm1.btnSjDisSelect.disabled = false


	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------
		lRow = .vspdData.MaxRows

		i=0
		If lRow > 0 And intARow > 0 Then
			If intTRow<=0 Then
				ReDim lgStrPrevKeyM(intARow - 1)
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
			Else
				TmpArrPrevKey=lgStrPrevKeyM
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lgStrPrevKeyM(intTRow+intARow - 1)
				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
				For i = 0 To intTRow-1
					lgStrPrevKeyM(i) = TmpArrPrevKey(i)
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next
			End If

			For i = intTRow To intTRow+intARow-1
				lglngHiddenRows(i) = 0
			Next

			if lgIntFlgModeM = Parent.OPMD_CMODE then
			    If DbQuery2(1, False) = False Then	Exit Function
		    End If
		    lgIntFlgModeM = Parent.OPMD_UMODE
		    lgIntFlgMode = Parent.OPMD_UMODE
		End If
	End With
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtPO_NO.focus
	End If
	Set gActiveElement = document.activeElement
    DbQueryOk = true
End Function

'=======================================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	DbQuery2 = False
	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim strPO_COMPANY,strSO_COMPANY,strPO_NO

	Call LayerShowHide(1)


	With frm1
		.vspdData.redraw = false
		.vspdData.Row = Row

		.vspdData.Col = C_PO_COMPANY
		strPO_COMPANY  = .vspdData.Text

		.vspdData.Col = C_SO_COMPANY
		strSO_COMPANY  = .vspdData.Text

		.vspdData.Col = C_PO_NO
		strPO_NO      = .vspdData.Text

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows

		strVal = strVal & "&strPO_COMPANY=" & trim(strPO_COMPANY)
		strVal = strVal & "&strSO_COMPANY=" & trim(strSO_COMPANY)
		strVal = strVal & "&strPO_NO=" & trim(strPO_NO)

		strVal = strVal & "&lgStrPrevKeyM="  & lgStrPrevKeyM(Row - 1)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'☜: Next key tag
		strVal = strVal & "&lglngHiddenRows=" & .vspdData.MaxRows


		.vspdData.redraw = True

	End With


	Call RunMyBizASP(MyBizASP, strVal)
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk2(Byval DataCount)
	DbQueryOk2 = false
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim Index

	With frm1.vspdData2
		lngRangeFrom = .MaxRows - DataCount + 1
		lngRangeTo = .MaxRows
	End With

	DbQueryOk2 = true

End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave()
	Dim lRow
	Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel, strTxt, k

	DbSave = False

	Call LayerShowHide(1)
	'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
	.txtMode.value = parent.UID_M0002

	<%  '-----------------------
	'Data manipulate area
	'----------------------- %>
	lGrpCnt = 1

	strVal = ""
	strDel = ""

	<%  '-----------------------
	'Data manipulate area
	'----------------------- %>
	' Data 연결 규칙 
	' 0: Flag , 1: Row위치, 2~N: 각 데이타 


	For lRow = 1 To .vspdData.MaxRows

		.vspdData.Row = lRow
		.vspdData.Col = 0

		Select Case .vspdData.Text

		    Case ggoSpread.InsertFlag											'☜: 신규 

					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep				'☜: C=Create
		                        strVal = strVal & parent.gRowSep

		        		lGrpCnt = lGrpCnt + 1

		    Case ggoSpread.UpdateFlag											'☜: 신규 

					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep				'☜: U=Update
					'상단 스프레드 
					.vspdData.Col =C_CFM_YN		 	'확정여부 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col =C_PO_COMPANY    		'발주법인 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col =C_SO_COMPANY       	'수주법인 
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col =C_PO_NO              	'발주번호(고객주문번호)
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep


					'수주형태 	txtSo_Type
					strVal = strVal & Trim(.txtSo_Type.value) & parent.gColSep
					'수주일		txtSo_dt
					strVal = strVal & Trim(.txtSo_dt.Text) & parent.gColSep
					'판매유형	txtDeal_Type
					strVal = strVal & Trim(.txtDeal_Type.value) & parent.gColSep
					'영업그룹	txtSales_Grp
					strVal = strVal & Trim(.txtSales_Grp.value) & parent.gColSep
					'공장	txtPlantCd
					strVal = strVal & Trim(.txtPlantCd.value) & parent.gColSep
					'유효일 txtvalid_dt
					strVal = strVal & Trim(.txtvalid_dt.Text) & parent.gColSep

		                        strVal = strVal & parent.gRowSep

		        		lGrpCnt = lGrpCnt + 1

		    Case ggoSpread.DeleteFlag													'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep						'☜: U=Update
		                        strVal = strVal & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1
		End Select

    	Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>

	End With


    DbSave = True
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call MainQuery()


End Function



'------------------------------------------  OpenSupplier()  -------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주법인"
	arrParam(1) = "B_Biz_Partner"

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
'	arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE In ('C','CS') And IN_OUT_FLAG = 'O'"
	arrParam(5) = "발주법인"

    	arrField(0) = "BP_Cd"
    	arrField(1) = "BP_NM"

    	arrHeader(0) = "발주법인"
    	arrHeader(1) = "발주법인명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)
		frm1.txtSupplierNm.Value    = arrRet(1)
		frm1.txtSupplierCd.focus
		lgBlnFlgChgValue = True
	End If
End Function


'==========================================================================================================
Function OpenRequried(ByVal iRequried)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strpostFlag

	If IsOpenPop = True Then Exit Function
	If lsClickCfmYes = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried
	Case 0
		If lsClickCfmNo = True Then
			IsOpenPop = False
			Exit Function
		End If

		If UCase(frm1.txtSo_Type.className) = parent.UCN_PROTECTED Then
			IsOpenPop = False
			Exit Function
		End If


		If frm1.rdoPostFlag2(1).checked = true Then
			strpostFlag=" AND EXPORT_FLAG = 'Y' "			'외자이면 
		Else
			strpostFlag=" AND EXPORT_FLAG <> 'Y' "		'내자이면 
		End If


		arrParam(0) = "수주형태"
		arrParam(1) = "S_SO_TYPE_CONFIG"
		arrParam(2) = Trim(frm1.txtSo_Type.value)
		arrParam(3) = Trim(frm1.txtSo_TypeNm.value)
		arrParam(4) = "intercom_flg = 'Y' " & strpostFlag & " "
		arrParam(5) = "수주형태"

		arrField(0) = "SO_TYPE"
	    arrField(1) = "SO_TYPE_NM"
	    arrField(2) = "EXPORT_FLAG"
	    arrField(3) = "RET_ITEM_FLAG"
	    arrField(4) = "AUTO_DN_FLAG"
	    arrField(5) = "CI_FLAG"

	    arrHeader(0) = "수주형태"
	    arrHeader(1) = "수주형태명"
	    arrHeader(2) = "수출여부"
	    arrHeader(3) = "반품여부"
	    arrHeader(4) = "자동출하생성여부"
	    arrHeader(5) = "통관여부"

		frm1.txtSo_Type.focus


	Case 1
		If lsClickCfmNo = True Then
			IsOpenPop = False
			Exit Function
		End If

		If UCase(frm1.txtSales_Grp.className) = parent.UCN_PROTECTED Then
			IsOpenPop = False
			Exit Function
		End IF

		arrParam(0) = "영업그룹"
		arrParam(1) = "B_SALES_GRP"
		arrParam(2) = Trim(frm1.txtSales_Grp.value)
		arrParam(3) = Trim(frm1.txtSales_GrpNm.value)
		arrParam(4) = "USAGE_FLAG='Y'"
		arrParam(5) = "영업그룹"

	    arrField(0) = "SALES_GRP"
	    arrField(1) = "SALES_GRP_NM"

	    arrHeader(0) = "영업그룹"
	    arrHeader(1) = "영업그룹명"

		frm1.txtSales_Grp.focus
	End Select

	arrParam(3) = ""			'☜: [Condition Name Delete]

	If iRequried = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	End If

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRequried(arrRet,iRequried)
	End If
End Function


Function OpenOption(ByVal iOption)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iOption
	Case 1
		If frm1.txtDeal_Type.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If

		If lsClickCfmYes = True or lsClickCfmNo = True Then
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = "판매유형"
		arrParam(1) = "B_MINOR"
		arrParam(2) = Trim(frm1.txtDeal_Type.value)
		arrParam(3) = Trim(frm1.txtDeal_Type_nm.value)
		arrParam(4) = "MAJOR_CD='S0001'"
		arrParam(5) = "판매유형"

		arrField(0) = "MINOR_CD"
		arrField(1) = "MINOR_NM"

		arrHeader(0) = "판매유형"
		arrHeader(1) = "판매유형명"

		frm1.txtDeal_Type.focus
	End Select

	arrParam(3) = ""			'☜: [Condition Name Delete]

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtDeal_Type.value = arrRet(0)
		frm1.txtDeal_Type_nm.value = arrRet(1)
	End If
End Function

'============================================================================================================
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"

	arrField(0) = "Plant_cd"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)

	End If

End Function

'==========================================================================================================
Function SetRequried(Byval arrRet,ByVal iRequried)

	If arrRet(0) <> "" Then

		Select Case iRequried
		Case 0
			frm1.txtSo_Type.value = arrRet(0)
			frm1.txtSo_TypeNm.value = arrRet(1)
		Case 1
			frm1.txtSales_Grp.value = arrRet(0)
			frm1.txtSales_GrpNm.value = arrRet(1)
		End Select

		lgBlnFlgChgValue = True

	End If

End Function


'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 일괄선택 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CfmFlg
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 0 then
			    frm1.vspdData.value = 1
                Call vspdData_ButtonClicked(C_CfmFlg, i, 1)
		    end if
		Next
	End If
End Sub


'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 일괄선택취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CfmFlg
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 1 then
			    frm1.vspdData.value = 0
                Call vspdData_ButtonClicked(C_CfmFlg, i, 0)
		    end if
		Next
	End If
End Sub



'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 일괄확정 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnSjSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CFM_YN
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 0 then
			    frm1.vspdData.value = 1
                Call vspdData_ButtonClicked(C_CFM_YN, i, 1)
		    end if
		Next
	End If

	Call btnSelect_OnClick()
End Sub


'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 일괄확정취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnSjDisSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CFM_YN
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 1 then
			    frm1.vspdData.value = 0
                Call vspdData_ButtonClicked(C_CFM_YN, i, 0)
		    end if
		Next
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<!--########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################## -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>멀티컴퍼니 수주등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;<label id="lblT" name="lblTest"></label></TD>
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
								<TD CLASS="TD5" NOWRAP>발주법인</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtSupplierCd"  SIZE=10 MAXLENGTH=10 ALT="발주법인"  tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
										       <INPUT TYPE=TEXT Name="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="발주법인"  tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>고객발주일</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/ksm111ma1_fpDateTime2_txtFrDt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/ksm111ma1_fpDateTime2_txtToDt.js'></script>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>내외자처리구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoPostFlag2" id="rdoPostFlagN" value="N" tag = "11" checked>
										<label for="rdoPostFlagN">내자</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoPostFlag2" id="rdoPostFlagY" value="Y" tag = "11" >
										<label for="rdoPostFlagY">외자</label>
								</TD>
								<TD CLASS="TD5" NOWRAP>고객발주번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtPO_NO" SIZE=29 MAXLENGTH=18 tag="11" ALT="고객발주번호" STYLE="text-transform:uppercase"></TD>
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
				  <TABLE <%=LR_SPACE_TYPE_60%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>수주형태</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_Type" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 0" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">
								     <INPUT NAME="txtSo_TypeNm" TYPE="Text" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
						<TD CLASS=TD5 NOWRAP>수주일</TD>
						<TD CLASS=TD6 NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
									<script language =javascript src='./js/ksm111ma1_fpDateTime1_txtSo_dt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>판매유형</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeal_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 1">
						                     <INPUT NAME="txtDeal_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">
								     <INPUT NAME="txtSales_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP ><INPUT NAME="txtPlantCd" ALT="공장" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConPlant()">
								       <INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>유효일</TD>
						<TD CLASS=TD6 NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
									<script language =javascript src='./js/ksm111ma1_fpDateTime1_txtvalid_dt.js'></script>
									</td>
								<tr>
							</table>
						</TD>

					</TR>
					<TR HEIGHT=70%>
						<TD WIDTH=100% COLSPAN=4>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/ksm111ma1_A_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR HEIGHT= 30%>
						<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						     <script language =javascript src='./js/ksm111ma1_B_vspdData2.js'></script>
						</TD>
					</TR>
				  </TABLE>
				 </TD>
			</TR>
		</TABLE>

		</TD>
	</TR>

	<TR HEIGHT="20">
	<TD WIDTH="100%">
		<table  CLASS="BasicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<td WIDTH="*" align="left">
				<BUTTON name="btnSelect" class="clsmbtn" >일괄선택</button>&nbsp;
				<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>&nbsp;&nbsp;
				<BUTTON name="btnSjSelect" class="clsmbtn" >일괄확정</button>&nbsp;
				<BUTTON NAME="btnSjDisSelect" CLASS="CLSMBTN">일괄확정취소</BUTTON>
				</TD>
				<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">멀티컴퍼니수주조회</a></td>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
	</TD>
	</TR>
	 <TR>
	  <TD WIDTH=100% HEIGHT="<%=BizSize%>"><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT="<%=BizSize%>" FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex = -1></IFRAME>
	  </TD>
	 </TR>
</TABLE>

<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex = -1></TEXTAREA>
<Input TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<Input TYPE=HIDDEN NAME="txtFlgMode" tag="24">


<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnState" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMrp" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPrTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
