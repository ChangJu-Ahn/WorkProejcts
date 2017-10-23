<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111qa1
'*  4. Program Name         : 멀티컴퍼니수주조회 
'*  5. Program Desc         : 멀티컴퍼니수주조회-멀티 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "ksm112mb1.asp"
Const BIZ_PGM_ID2 = "ksm112mb01.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL = "S3111MA1"
											'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'상단 스프레드 
Dim C_SEL_YN 	         		'선택 
Dim C_CFM_FLAG				'수주확정여부 
Dim C_SOLD_TO_PARTY			'발주법인 
Dim C_BP_FULL_NM			'발주법인명 
Dim C_CUST_PO_NO			'고객발주번호 
Dim C_SO_NO				'수주번호 
Dim C_EXPORT_FLAG			'내외자구분 
Dim C_SO_DT				'수주일 
Dim C_SALES_GRP				'영업그룹 
Dim C_SALES_GRP_FULL_NM			'영업그룹명 
Dim C_CUR				'화페 
Dim C_NET_AMT				'수주금액 
Dim C_VAT_AMT				'부가세금액 
Dim C_NET_VAT_TOTAMT			'수주총금액 
Dim C_VAT_TYPE				'부가세유형 
Dim C_VAT_TYPE_NM			'부가세유형명 
Dim C_VAT_RATE				'부가세율 
Dim C_PAY_METH				'결제방법 
Dim C_PAY_METH_NM			'결제방법명 
Dim C_INCOTERMS				'가격조건 
Dim C_INCOTERMS_NM			'가격조건명 
Dim C_HIDDEN_CFM_FLAG			'수주확정여부(HIDDEN)

'하단 스프레드 
Dim C_ITEM_CD				'품목 
Dim C_ITEM_NM				'품목명 
Dim C_SPEC				'품목규격 
Dim C_CUST_ITEM_CD			'고객품목 
Dim C_BP_ITEM_NM			'고객품목명 
Dim C_BP_ITEM_SPEC			'고객품목규격 
Dim C_SO_QTY				'수량 
Dim C_SO_UNIT				'단위 
Dim C_SO_PRICE				'단가 
Dim C_NET_AMT2				'금액 
Dim C_DLVY_DT				'납기일 
Dim C_VAT_AMT2				'부가세금액 
Dim C_VAT_RATE2				'부가세율 
Dim C_VAT_TYPE2				'부가세유형 
Dim C_VAT_TYPE_NM2			'부가세유형명 
Dim C_VAT_INC_FLAG			'부가세포함구분 



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
Dim EndDate, StartDate,CurrDate, iDBSYSDate,iBoDate
iDBSYSDate = "<%=GetSvrDate%>"
CurrDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate = UnIDateAdd("m", 1, CurrDate, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, CurrDate, parent.gDateFormat)
iBoDate = UnIDateAdd("d", -15, CurrDate, parent.gDateFormat)

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

	frm1.txtSo_Frdt.Text = iBoDate
	frm1.txtSo_Todt.Text = CurrDate

	Call SetToolbar("1100000000001111")



	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True

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

		.MaxCols = C_HIDDEN_CFM_FLAG + 1
		.Col = .MaxCols:	.ColHidden = True
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck    C_SEL_YN, "선택",10,,,true
		ggoSpread.SSSetEdit 	C_CFM_FLAG		,"수주확정여부",20
		ggoSpread.SSSetEdit 	C_SOLD_TO_PARTY		,"발주법인"		,15		'발주법인 
		ggoSpread.SSSetEdit 	C_BP_FULL_NM		,"발주법인명"		,20		'발주법인명 
		ggoSpread.SSSetEdit 	C_CUST_PO_NO		,"고객발주번호"	,15		'고객발주번호 
		ggoSpread.SSSetEdit 	C_SO_NO			,"수주번호"		,15		'수주번호 
		ggoSpread.SSSetEdit 	C_EXPORT_FLAG		,"내외자구분"		,15		'내외자구분 
		ggoSpread.SSSetDate 	C_SO_DT			,"수주일"		,		10,		2,					parent.gDateFormat'수주일 
		ggoSpread.SSSetEdit 	C_SALES_GRP		,"영업그룹"		,20		'영업그룹 
		ggoSpread.SSSetEdit 	C_SALES_GRP_FULL_NM	,"영업그룹명"		,20		'영업그룹명 
		ggoSpread.SSSetEdit 	C_CUR			,"화페"			,20		'화페 
		ggoSpread.SSSetFloat 	C_NET_AMT		,"수주금액"			,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_VAT_AMT		,"부가세금액"		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_NET_VAT_TOTAMT	,"수주총금액"		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_VAT_TYPE		,"부가세유형"		,		10,		,					,	  5,	  2
		ggoSpread.SSSetEdit 	C_VAT_TYPE_NM		,"부가세유형명"		,20		'부가세유형명 
		ggoSpread.SSSetFloat 	C_VAT_RATE		,"부가세율"		,		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_PAY_METH		,"결제방법"		,20		'결제방법 
		ggoSpread.SSSetEdit 	C_PAY_METH_NM		,"결제방법명"		,20		'결제방법명 
		ggoSpread.SSSetEdit 	C_INCOTERMS		,"가격조건"		,20		'가격조건 
		ggoSpread.SSSetEdit 	C_INCOTERMS_NM		,"가격조건명"		,20		'가격조건명 
		ggoSpread.SSSetEdit 	C_HIDDEN_CFM_FLAG	,"수주확정여부"		,10		'수주확정여부(HIDDEN)


		Call ggoSpread.SSSetColHidden(C_HIDDEN_CFM_FLAG,	C_HIDDEN_CFM_FLAG,	True)
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

		.MaxCols = C_VAT_INC_FLAG+1
		.Col = .MaxCols:	.ColHidden = True

		.MaxRows = 0

		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit 	C_ITEM_CD	,"품목" 		,			18,		,					,	  18,	  2                                                                                  					'품목 
		ggoSpread.SSSetEdit 	C_ITEM_NM	,"품목명" 		,		25,		,					,	  40                                                                                                 					'품목명 
		ggoSpread.SSSetEdit 	C_SPEC		,"품목규격"		,			20
		ggoSpread.SSSetEdit 	C_CUST_ITEM_CD	,"고객품목"		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_BP_ITEM_NM	,"고객품목명"	,		25,		,					,	  40
		ggoSpread.SSSetEdit 	C_BP_ITEM_SPEC	,"고객품목규격" 	,			20
		ggoSpread.SSSetFloat 	C_SO_QTY	,"수량" 		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_SO_UNIT	,"단위" 		,			8,		,					,	  3,	  2
		ggoSpread.SSSetFloat 	C_SO_PRICE	,"단가" 		,			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_NET_AMT2	,"금액" 		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetDate 	C_DLVY_DT	,"납기일" 		,		10,		2,					parent.gDateFormat
		ggoSpread.SSSetFloat 	C_VAT_AMT2	,"부가세금액" 	,		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_VAT_RATE2	,"부가세율"		,		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_VAT_TYPE2	,"부가세유형"	,		10,		,					,	  5,	  2
		ggoSpread.SSSetEdit 	C_VAT_TYPE_NM2	,"부가세유형명" 	,	20
		ggoSpread.SSSetEdit 	C_VAT_INC_FLAG	,"부가세포함구분" 	,20		'부가세포함구분 

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

		ggoSpread.SpreadLock 	C_CFM_FLAG		, -1, -1		'수주확정여부 
		ggoSpread.SpreadLock 	C_SOLD_TO_PARTY		, -1, -1		'발주법인 
		ggoSpread.SpreadLock 	C_BP_FULL_NM		, -1, -1		'발주법인명 
		ggoSpread.SpreadLock 	C_CUST_PO_NO		, -1, -1		'고객발주번호 
		ggoSpread.SpreadLock 	C_SO_NO			, -1, -1		'수주번호 
		ggoSpread.SpreadLock 	C_EXPORT_FLAG		, -1, -1		'내외자구분 
		ggoSpread.SpreadLock 	C_SO_DT			, -1, -1		'수주일 
		ggoSpread.SpreadLock 	C_SALES_GRP		, -1, -1		'영업그룹 
		ggoSpread.SpreadLock 	C_SALES_GRP_FULL_NM	, -1, -1		'영업그룹명 
		ggoSpread.SpreadLock 	C_CUR			, -1, -1		'화페 
		ggoSpread.SpreadLock 	C_NET_AMT		, -1, -1		'수주금액 
		ggoSpread.SpreadLock 	C_VAT_AMT		, -1, -1		'부가세금액 
		ggoSpread.SpreadLock 	C_NET_VAT_TOTAMT	, -1, -1		'수주총금액 
		ggoSpread.SpreadLock 	C_VAT_TYPE		, -1, -1		'부가세유형 
		ggoSpread.SpreadLock 	C_VAT_TYPE_NM		, -1, -1		'부가세유형명 
		ggoSpread.SpreadLock 	C_VAT_RATE		, -1, -1		'부가세율 
		ggoSpread.SpreadLock 	C_PAY_METH		, -1, -1		'결제방법 
		ggoSpread.SpreadLock 	C_PAY_METH_NM		, -1, -1		'결제방법명 
		ggoSpread.SpreadLock 	C_INCOTERMS		, -1, -1		'가격조건 
		ggoSpread.SpreadLock 	C_INCOTERMS_NM		, -1, -1		'가격조건명 
		ggoSpread.SpreadLock 	C_HIDDEN_CFM_FLAG	, -1, -1	'수주확정여부(HIDDEN)


		.ReDraw = True
	End With
End Sub

Sub SetSpreadLock2()
	With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2

		ggoSpread.SpreadLock	C_ITEM_CD	, -1, -1	'품목 
		ggoSpread.SpreadLock	C_ITEM_NM	, -1, -1	'품목명 
		ggoSpread.SpreadLock	C_SPEC		, -1, -1	'품목규격 
		ggoSpread.SpreadLock	C_CUST_ITEM_CD	, -1, -1	'고객품목 
		ggoSpread.SpreadLock	C_BP_ITEM_NM	, -1, -1	'고객품목명 
		ggoSpread.SpreadLock	C_BP_ITEM_SPEC	, -1, -1	'고객품목규격 
		ggoSpread.SpreadLock	C_SO_QTY	, -1, -1	'수량 
		ggoSpread.SpreadLock	C_SO_UNIT	, -1, -1	'단위 
		ggoSpread.SpreadLock	C_SO_PRICE	, -1, -1	'단가 
		ggoSpread.SpreadLock	C_NET_AMT2	, -1, -1	'금액 
		ggoSpread.SpreadLock	C_DLVY_DT	, -1, -1	'납기일 
		ggoSpread.SpreadLock	C_VAT_AMT2	, -1, -1	'부가세금액 
		ggoSpread.SpreadLock	C_VAT_RATE2	, -1, -1	'부가세율 
		ggoSpread.SpreadLock	C_VAT_TYPE2	, -1, -1	'부가세유형 
		ggoSpread.SpreadLock	C_VAT_TYPE_NM2	, -1, -1	'부가세유형명 
		ggoSpread.SpreadLock	C_VAT_INC_FLAG	, -1, -1	'부가세포함구분 

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


		ggoSpread.SSSetProtected 	C_CFM_FLAG		, pvStartRow, pvEndRow		'수주확정여부		ggoSpread.SSSetProtected 	C_SOLD_TO_PARTY		, pvStartRow, pvEndRow		'발주법인 
		ggoSpread.SSSetProtected 	C_BP_FULL_NM		, pvStartRow, pvEndRow		'발주법인명 
		ggoSpread.SSSetProtected 	C_CUST_PO_NO		, pvStartRow, pvEndRow		'고객발주번호 
		ggoSpread.SSSetProtected 	C_SO_NO			, pvStartRow, pvEndRow		'수주번호 
		ggoSpread.SSSetProtected 	C_EXPORT_FLAG		, pvStartRow, pvEndRow		'내외자구분 
		ggoSpread.SSSetProtected 	C_SO_DT			, pvStartRow, pvEndRow		'수주일 
		ggoSpread.SSSetProtected 	C_SALES_GRP		, pvStartRow, pvEndRow		'영업그룹 
		ggoSpread.SSSetProtected 	C_SALES_GRP_FULL_NM	, pvStartRow, pvEndRow		'영업그룹명 
		ggoSpread.SSSetProtected 	C_CUR			, pvStartRow, pvEndRow		'화페 
		ggoSpread.SSSetProtected 	C_NET_AMT		, pvStartRow, pvEndRow		'수주금액 
		ggoSpread.SSSetProtected 	C_VAT_AMT		, pvStartRow, pvEndRow		'부가세금액 
		ggoSpread.SSSetProtected 	C_NET_VAT_TOTAMT	, pvStartRow, pvEndRow		'수주총금액 
		ggoSpread.SSSetProtected 	C_VAT_TYPE		, pvStartRow, pvEndRow		'부가세유형 
		ggoSpread.SSSetProtected 	C_VAT_TYPE_NM		, pvStartRow, pvEndRow		'부가세유형명 
		ggoSpread.SSSetProtected 	C_VAT_RATE		, pvStartRow, pvEndRow		'부가세율 
		ggoSpread.SSSetProtected 	C_PAY_METH		, pvStartRow, pvEndRow		'결제방법 
		ggoSpread.SSSetProtected 	C_PAY_METH_NM		, pvStartRow, pvEndRow		'결제방법명 
		ggoSpread.SSSetProtected 	C_INCOTERMS		, pvStartRow, pvEndRow		'가격조건 
		ggoSpread.SSSetProtected 	C_INCOTERMS_NM		, pvStartRow, pvEndRow		'가격조건명 
		ggoSpread.SSSetProtected 	C_HIDDEN_CFM_FLAG	, pvStartRow, pvEndRow		'수주확정여부(HIDDEN)

		.ReDraw = True
	End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
	With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2



		ggoSpread.SSSetProtected	C_ITEM_CD	, pvStartRow, pvEndRow	'품목 
		ggoSpread.SSSetProtected	C_ITEM_NM	, pvStartRow, pvEndRow	'품목명 
		ggoSpread.SSSetProtected	C_SPEC		, pvStartRow, pvEndRow	'품목규격 
		ggoSpread.SSSetProtected	C_CUST_ITEM_CD	, pvStartRow, pvEndRow	'고객품목 
		ggoSpread.SSSetProtected	C_BP_ITEM_NM	, pvStartRow, pvEndRow	'고객품목명 
		ggoSpread.SSSetProtected	C_BP_ITEM_SPEC	, pvStartRow, pvEndRow	'고객품목규격 
		ggoSpread.SSSetProtected	C_SO_QTY	, pvStartRow, pvEndRow	'수량 
		ggoSpread.SSSetProtected	C_SO_UNIT	, pvStartRow, pvEndRow	'단위 
		ggoSpread.SSSetProtected	C_SO_PRICE	, pvStartRow, pvEndRow	'단가 
		ggoSpread.SSSetProtected	C_NET_AMT2	, pvStartRow, pvEndRow	'금액 
		ggoSpread.SSSetProtected	C_DLVY_DT	, pvStartRow, pvEndRow	'납기일 
		ggoSpread.SSSetProtected	C_VAT_AMT2	, pvStartRow, pvEndRow	'부가세금액 
		ggoSpread.SSSetProtected	C_VAT_RATE2	, pvStartRow, pvEndRow	'부가세율 
		ggoSpread.SSSetProtected	C_VAT_TYPE2	, pvStartRow, pvEndRow	'부가세유형 
		ggoSpread.SSSetProtected	C_VAT_TYPE_NM2	, pvStartRow, pvEndRow	'부가세유형명 
		ggoSpread.SSSetProtected	C_VAT_INC_FLAG	, pvStartRow, pvEndRow	'부가세포함구분 

		.ReDraw = True
	End With
End Sub

'============================= 2.2.3 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_SEL_YN		= 1		'선택 
	C_CFM_FLAG		= 2		'수주확정여부 
	C_SOLD_TO_PARTY		= 3		'발주법인 
	C_BP_FULL_NM		= 4		'발주법인명 
	C_CUST_PO_NO		= 5		'고객발주번호 
	C_SO_NO			= 6		'수주번호 
	C_EXPORT_FLAG		= 7		'내외자구분 
	C_SO_DT			= 8		'수주일 
	C_SALES_GRP		= 9		'영업그룹 
	C_SALES_GRP_FULL_NM	= 10		'영업그룹명 
	C_CUR			= 11		'화페 
	C_NET_AMT		= 12		'수주금액 
	C_VAT_AMT		= 13		'부가세금액 
	C_NET_VAT_TOTAMT	= 14		'수주총금액 
	C_VAT_TYPE		= 15		'부가세유형 
	C_VAT_TYPE_NM		= 16		'부가세유형명 
	C_VAT_RATE		= 17		'부가세율 
	C_PAY_METH		= 18		'결제방법 
	C_PAY_METH_NM		= 19		'결제방법명 
	C_INCOTERMS		= 20		'가격조건 
	C_INCOTERMS_NM		= 21		'가격조건명 
	C_HIDDEN_CFM_FLAG	= 22		'수주확정여부(HIDDEN)
End Sub

Sub InitSpreadPosVariables2()
	C_ITEM_CD	= 1	'품목 
	C_ITEM_NM	= 2	'품목명 
	C_SPEC		= 3	'품목규격 
	C_CUST_ITEM_CD	= 4	'고객품목 
	C_BP_ITEM_NM	= 5	'고객품목명 
	C_BP_ITEM_SPEC	= 6	'고객품목규격 
	C_SO_QTY	= 7	'수량 
	C_SO_UNIT	= 8	'단위 
	C_SO_PRICE	= 9	'단가 
	C_NET_AMT2	= 10	'금액 
	C_DLVY_DT	= 11	'납기일 
	C_VAT_AMT2	= 12	'부가세금액 
	C_VAT_RATE2	= 13	'부가세율 
	C_VAT_TYPE2	= 14	'부가세유형 
	C_VAT_TYPE_NM2	= 15	'부가세유형명 
	C_VAT_INC_FLAG	= 16	'부가세포함구분 
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
			C_SEL_YN		= iCurColumnPos(1)	'선택 
			C_CFM_FLAG		= iCurColumnPos(2)	'수주확정여부 
			C_SOLD_TO_PARTY		= iCurColumnPos(3)	'발주법인 
			C_BP_FULL_NM		= iCurColumnPos(4)	'발주법인명 
			C_CUST_PO_NO		= iCurColumnPos(5)	'고객발주번호 
			C_SO_NO			= iCurColumnPos(6)	'수주번호 
			C_EXPORT_FLAG		= iCurColumnPos(7)	'내외자구분 
			C_SO_DT			= iCurColumnPos(8)	'수주일 
			C_SALES_GRP		= iCurColumnPos(9)	'영업그룹 
			C_SALES_GRP_FULL_NM	= iCurColumnPos(10)	'영업그룹명 
			C_CUR			= iCurColumnPos(11)	'화페 
			C_NET_AMT		= iCurColumnPos(12)	'수주금액 
			C_VAT_AMT		= iCurColumnPos(13)	'부가세금액 
			C_NET_VAT_TOTAMT	= iCurColumnPos(14)	'수주총금액 
			C_VAT_TYPE		= iCurColumnPos(15)	'부가세유형 
			C_VAT_TYPE_NM		= iCurColumnPos(16)	'부가세유형명 
			C_VAT_RATE		= iCurColumnPos(17)	'부가세율 
			C_PAY_METH		= iCurColumnPos(18)	'결제방법 
			C_PAY_METH_NM		= iCurColumnPos(19)	'결제방법명 
			C_INCOTERMS		= iCurColumnPos(20)	'가격조건 
			C_INCOTERMS_NM		= iCurColumnPos(21)	'가격조건명 
			C_HIDDEN_CFM_FLAG	= iCurColumnPos(22)	'수주확정여부(HIDDEN)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
            		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)


			C_ITEM_CD	= iCurColumnPos(1)	'품목 
			C_ITEM_NM	= iCurColumnPos(2)	'품목명 
			C_SPEC		= iCurColumnPos(3)	'품목규격 
			C_CUST_ITEM_CD	= iCurColumnPos(4)	'고객품목 
			C_BP_ITEM_NM	= iCurColumnPos(5)	'고객품목명 
			C_BP_ITEM_SPEC	= iCurColumnPos(6)	'고객품목규격 
			C_SO_QTY	= iCurColumnPos(7)	'수량 
			C_SO_UNIT	= iCurColumnPos(8)	'단위 
			C_SO_PRICE	= iCurColumnPos(9)	'단가 
			C_NET_AMT2	= iCurColumnPos(10)	'금액 
			C_DLVY_DT	= iCurColumnPos(11)	'납기일 
			C_VAT_AMT2	= iCurColumnPos(12)	'부가세금액 
			C_VAT_RATE2	= iCurColumnPos(13)	'부가세율 
			C_VAT_TYPE2	= iCurColumnPos(14)	'부가세유형 
			C_VAT_TYPE_NM2	= iCurColumnPos(15)	'부가세유형명 
			C_VAT_INC_FLAG	= iCurColumnPos(16)	'부가세포함구분 

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


'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD

	Dim istrSO_NO

	On Error Resume Next

	Const CookieSplit = 4877


	If Kubun = 0 Then

		If ReadCookie("txtSoCompanyCd") <> "" Then
			frm1.txtSupplierCd.Value = ReadCookie("txtSoCompanyCd")
		End If

		If ReadCookie("txtFrDt") <> "" Then
			frm1.txtFrDt.text = ReadCookie("txtFrDt")
		End If

		If ReadCookie("txtToDt") <> "" Then
			frm1.txtToDt.text = ReadCookie("txtToDt")
		End If

		If ReadCookie("txtSoCompanyCd") <> "" Then
			Call MainQuery()
		End If

		WriteCookie "txtSoCompanyCd", ""
		WriteCookie "txtFrDt", ""
		WriteCookie "txtToDt", ""

	elseIf Kubun = 1 Then

	    If lgIntFlgMode <> Parent.OPMD_UMODE Then
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If

	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If


	With frm1.vspdData
		.Row		= .ActiveRow
		.Col		= C_SO_NO
		istrSO_NO	= Trim(.text)
	End With

	WriteCookie CookieSplit , istrSO_NO

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

    	lRow = .vspdData2.ActiveRow
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

Sub rdoCfmFlagN_onClick()
'	frm1.vspdData.MaxRows = 0
'	frm1.vspdData2.MaxRows = 0
'	frm1.btnSelect.disabled = True
'	frm1.btnDisSelect.disabled = True
'	Call fncquery()
End Sub

Sub rdoCfmFlagY_onClick()
'	frm1.vspdData.MaxRows = 0
'	frm1.vspdData2.MaxRows = 0
'	frm1.btnSelect.disabled = True
'	frm1.btnDisSelect.disabled = True
'	Call fncquery()
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
	If Col = C_SEL_YN And Row > 0 Then
		frm1.vspdData.Redraw = false

		.Col = C_SEL_YN
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
'   Event Name : txtSo_Frdt
'   Event Desc : 수주일 
'==========================================================================================
 Sub txtSo_Todt_DblClick(Button)
	if Button = 1 then
		frm1.txtSo_Frdt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtSo_Frdt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : txtSo_Todt
'   Event Desc : 수주일 
'==========================================================================================
 Sub txtSo_Tordt_DblClick(Button)
	if Button = 1 then
		frm1.txtSo_Todt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtSo_Todt.Focus
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

Sub txtSo_Frdt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtSo_Todt_KeyDown(KeyCode, Shift)
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
	End If																		'☜: Query db data

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
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'수주일 From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'발주일 To

			if .rdoCfmFlag(0).checked = true Then					'수주확정여부 
				strVal = strVal & "&rdoCfmFlag=" & "Y"	'확정 
			else
				strVal = strVal & "&rdoCfmFlag=" & "N"	'미확정 
			End if

			strVal = strVal & "&txtPO_NO=" & Trim(.txtPO_NO.value)			'고객발주번호 
			strVal = strVal & "&txtSO_NO=" & Trim(.txtSO_NO.value)			'수주번호 

			strVal = strVal & "&lgPageNo=" & lgPageNo                  		'☜: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	Else
	        	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '발주법인코드 
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)			'고객발주일 From
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)			'고객발주일 To
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'수주일 From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'발주일 To

			if .rdoCfmFlag(0).checked = true Then					'수주확정여부 
				strVal = strVal & "&rdoCfmFlag=" & "Y"	'확정 
			else
				strVal = strVal & "&rdoCfmFlag=" & "N"	'미확정 
			End if

			strVal = strVal & "&txtPO_NO=" & Trim(.txtPO_NO.value)			'고객발주번호 
			strVal = strVal & "&txtSO_NO=" & Trim(.txtSO_NO.value)			'수주번호 

			strVal = strVal & "&lgPageNo=" & lgPageNo                  		'☜: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	End If
	End with



	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 

	DbQuery = True
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

	if frm1.rdoCfmFlag(0).checked = true Then	'확정상태라면 
		Call SetToolBar("11001000000011")				'버튼 툴바 제어 
	Else						'미확정상태라면 
		Call SetToolBar("11001011000011")				'버튼 툴바 제어 
	End If

	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False



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
	Dim strSO_NO

	Call LayerShowHide(1)


	With frm1
		.vspdData.redraw = false
		.vspdData.Row = Row

		.vspdData.Col = C_SO_NO		'수주번호 
		strSO_NO  = .vspdData.Text

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows

		strVal = strVal & "&strSO_NO=" & trim(strSO_NO)

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
	Dim istrSO_NO
	Dim strHIDDEN_CFM_FLAG

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
		    Case ggoSpread.DeleteFlag													'☜: 삭제 
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep
					.vspdData.Row		= lRow
					.vspdData.Col		= C_SO_NO
					istrSO_NO	= Trim(.vspdData.text)
					strDel = strDel & Trim(istrSO_NO)		& parent.gColSep	'수주번호 

					.vspdData.Col		= C_HIDDEN_CFM_FLAG
					strHIDDEN_CFM_FLAG	= Trim(.vspdData.text)
					strDel = strDel & Trim(strHIDDEN_CFM_FLAG)		& parent.gColSep	'수주번호 
										'☜: U=Update
		                        strDel = strDel & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1

		    Case ggoSpread.UpdateFlag													'☜: 삭제 
					strDel = strDel & "U" & parent.gColSep & lRow & parent.gColSep
					.vspdData.Row		= lRow
					.vspdData.Col		= C_SO_NO
					istrSO_NO	= Trim(.vspdData.text)
					strDel = strDel & Trim(istrSO_NO)		& parent.gColSep	'수주번호 

					.vspdData.Col		= C_HIDDEN_CFM_FLAG
					strHIDDEN_CFM_FLAG	= Trim(.vspdData.text)
					strDel = strDel & Trim(strHIDDEN_CFM_FLAG)		& parent.gColSep	'수주번호 
										'☜: U=Update
		                        strDel = strDel & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1

		End Select

    	Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel

	If Trim(strDel)="" Then
		Call LayerShowHide(0)
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        	Exit Function
	End If


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>

	End With




    DbSave = True
End Function

'========================================================================================================
Function FncDeleteRow()
	Dim lDelRows, lDelRow

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if

    With Frm1.vspdData
    	.focus
    	ggoSpread.Source = frm1.vspdData
    	lDelRows = ggoSpread.DeleteRow

    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------


	frm1.vspdData.Col = C_SEL_YN
	frm1.vspdData.Row = .ActiveRow
	if frm1.vspdData.value = 0 then
		frm1.vspdData.value = 1
		Call vspdData_ButtonClicked(C_SEL_YN, .ActiveRow, 1)
	end if


    lgBlnFlgChgValue = True
    'Call TotalSum

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
Function FncDelete()
    Dim intRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '필요없는지???
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    If DbDelete = False Then                                                '☜: Delete db data
       Exit Function                                                        '☜:
    End If

    Call ggoOper.ClearField(Document, "A")

    '------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

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
			frm1.vspdData.Col = C_SEL_YN
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 0 then
			    frm1.vspdData.value = 1
                Call vspdData_ButtonClicked(C_SEL_YN, i, 1)
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
			frm1.vspdData.Col = C_SEL_YN
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 1 then
			    frm1.vspdData.value = 0
                Call vspdData_ButtonClicked(C_SEL_YN, i, 0)
		    end if
		Next
	End If
End Sub


'==========================================================================================
'   Event Name : btnProcessCfm_OnClick()
'   Event Desc : 확정처리 또는 확정취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnProcessCfm_OnClick()
	Dim lRow
	Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel, strTxt, k
	Dim istrSO_NO
	Dim strHIDDEN_CFM_FLAG


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
		    Case ggoSpread.UpdateFlag													'☜: 삭제 
					strDel = strDel & "U" & parent.gColSep & lRow & parent.gColSep
					.vspdData.Row		= lRow
					.vspdData.Col		= C_SO_NO
					istrSO_NO	= Trim(.vspdData.text)
					strDel = strDel & Trim(istrSO_NO)		& parent.gColSep	'수주번호 

					.vspdData.Col		= C_HIDDEN_CFM_FLAG
					strHIDDEN_CFM_FLAG	= Trim(.vspdData.text)
					strDel = strDel & Trim(strHIDDEN_CFM_FLAG)		& parent.gColSep	'수주번호 
										'☜: U=Update
		                        strDel = strDel & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1
		End Select

    	Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel

	If Trim(strDel)="" Then
		Call LayerShowHide(0)
		Call DisplayMsgBox("181216", "X", "X", "X")                          <%'No data changed!!%>
        	Exit Sub
	End If


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>

	End With

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>멀티컴퍼니 수주확정/삭제</font></td>
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
												<script language =javascript src='./js/ksm112ma1_fpDateTime2_txtFrDt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/ksm112ma1_fpDateTime2_txtToDt.js'></script>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>수주일</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/ksm112ma1_fpDateTime2_txtSo_Frdt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/ksm112ma1_fpDateTime2_txtSo_Todt.js'></script>
											</td>
										<tr>
									</table>
								</TD>
								<TD CLASS="TD5" NOWRAP>수주확정처리여부</TD>
								<TD CLASS="TD6" NOWRAP>
									<input type=radio CLASS = "RADIO" name="rdoCfmFlag" id="rdoCfmFlagN" value="Y" tag = "11" checked>
										<label for="rdoCfmFlagN">확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS="RADIO" name="rdoCfmFlag" id="rdoCfmFlagY" value="N" tag = "11" >
										<label for="rdoCfmFlagY">미확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>고객발주번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtPO_NO" SIZE=29 MAXLENGTH=18  tag="11" ALT="고객발주번호" STYLE="text-transform:uppercase"></TD>
								<TD CLASS="TD5" NOWRAP>수주번호</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtSO_NO" SIZE=29 MAXLENGTH=18 tag="11" ALT="수주번호" STYLE="text-transform:uppercase"></TD>

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
					<TR HEIGHT=60%>
						<TD WIDTH=100% COLSPAN=4>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/ksm112ma1_A_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR HEIGHT= 40%>
						<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						     <script language =javascript src='./js/ksm112ma1_B_vspdData2.js'></script>
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
				<TD WIDTH="*" align="left">
				<BUTTON name="btnSelect" class="clsmbtn" >일괄선택</button>&nbsp;
				<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>&nbsp;&nbsp;
				</TD>
				<td WIDTH="*" align="right"></td>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
	</TD>
	</TR>
	 <TR>
	  <TD WIDTH=100% HEIGHT=20><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex = -1></IFRAME>
	  </TD>
	 </TR>
</TABLE>

<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex = -1></TEXTAREA>
<Input TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
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
