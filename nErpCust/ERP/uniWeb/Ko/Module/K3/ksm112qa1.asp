<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm112qa1
'*  4. Program Name         : 멀티컴퍼니수발주진행조회(수주별)
'*  5. Program Desc         : 멀티컴퍼니수발주진행조회(수주별)-멀티 
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "ksm112qb1.asp"
											'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_PO_COMPANY	'발주법인 
Dim C_PO_COMPANY_NM	'발주법인명 
Dim C_SO_NO		'수주번호 
Dim C_SO_SEQ_NO		'수주순번 
Dim C_ITEM_CD		'품목 
Dim C_ITEM_NM		'품목명 
Dim C_SPEC		'품목규격 
Dim C_PO_STS		'발주법인상태 
Dim C_SO_STS		'수주법인상태 
Dim C_UNIT		'단위 
Dim C_PO_QTY		'발주수량 
Dim C_SO_QTY		'수주수량 
Dim C_PO_LC_QTY		'수입L/C수량 
Dim C_SO_LC_QTY		'수출L/C수량 
Dim C_SO_REQ_QTY	'출하요청수량 
Dim C_SO_ISSUE_QTY	'출고수량 
Dim C_SO_CC_QTY		'수출통관수량 
Dim C_PO_CC_QTY		'수입통관수량 
Dim C_PO_RCPT_QTY	'입고수량 
Dim C_SO_BILL_QTY	'매출수량 
Dim C_PO_IV_QTY		'매입수량 
Dim C_PO_NO		'고객주문번호 
Dim C_PO_SEQ_NO		'순번 
Dim C_BP_ITEM_CD	'고객품목 
Dim C_BP_ITEM_NM	'고객품목명 


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

Dim lgSortKey1

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

	lgLngCurRows = 0						'initializes Deleted Rows Count
	lgSortKey1 = 2
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
	frm1.txtSpplCd.value = parent.gCompany
	frm1.txtSpplNm.value = parent.gCompanyNm

	frm1.txtSo_Frdt.Text = iBoDate
	frm1.txtSo_Todt.Text = CurrDate

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

		.MaxCols = C_BP_ITEM_NM + 1
		.Col = .MaxCols:	.ColHidden = True
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_PO_COMPANY	        	,"발주법인"		,15	'발주법인 
		ggoSpread.SSSetEdit 	C_PO_COMPANY_NM	        	,"발주법인명"		,25	'발주법인 
		ggoSpread.SSSetEdit 	C_SO_NO		        	,"수주번호"		,15	'수주번호 
		ggoSpread.SSSetEdit 	C_SO_SEQ_NO			,"수주순번"		,15	'수주순번 
		ggoSpread.SSSetEdit 	C_ITEM_CD			,"품목"		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_ITEM_NM			,"품목명"		,		25,		,					,	  40
		ggoSpread.SSSetEdit 	C_SPEC		        	,"품목규격"		,			20
		ggoSpread.SSSetEdit 	C_PO_STS			,"발주법인상태"	,20	'발주법인상태 
		ggoSpread.SSSetEdit 	C_SO_STS			,"수주법인상태"	,20	'수주법인상태 
		ggoSpread.SSSetEdit 	C_UNIT		        	,"단위"		,			8,		,					,	  3,	  2
		ggoSpread.SSSetFloat 	C_PO_QTY			,"발주수량"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_QTY			,"수주수량"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_LC_QTY			,"수입L/C수량"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_LC_QTY			,"수출L/C수량"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_REQ_QTY	        	,"출하요청수량"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_ISSUE_QTY	        	,"출고수량"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_CC_QTY			,"수출통관수량"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_CC_QTY			,"수입통관수량"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_RCPT_QTY	        	,"입고수량"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_BILL_QTY	        	,"매출수량"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_IV_QTY			,"매입수량"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_PO_NO		        	,"고객주문번호"		,15	'고객주문번호 
		ggoSpread.SSSetEdit 	C_PO_SEQ_NO			,"순번"		,15	'순번 
		ggoSpread.SSSetEdit 	C_BP_ITEM_CD	        	,"고객품목"		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_BP_ITEM_NM	        	,"고객품목명"	,		25,		,					,	  40


		Call SetSpreadLock

	    	.ReDraw = true
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


		ggoSpread.SpreadLock 	C_PO_COMPANY	        , -1, -1	'발주법인 
		ggoSpread.SpreadLock 	C_PO_COMPANY_NM	        , -1, -1	'발주법인명 
		ggoSpread.SpreadLock 	C_SO_NO		        , -1, -1	'수주번호 
		ggoSpread.SpreadLock 	C_SO_SEQ_NO		, -1, -1	'수주순번 
		ggoSpread.SpreadLock 	C_ITEM_CD		, -1, -1	'품목 
		ggoSpread.SpreadLock 	C_ITEM_NM		, -1, -1	'품목명 
		ggoSpread.SpreadLock 	C_SPEC		        , -1, -1	'품목규격 
		ggoSpread.SpreadLock 	C_PO_STS		, -1, -1	'발주법인상태 
		ggoSpread.SpreadLock 	C_SO_STS		, -1, -1	'수주법인상태 
		ggoSpread.SpreadLock 	C_UNIT		        , -1, -1	'단위 
		ggoSpread.SpreadLock 	C_PO_QTY		, -1, -1	'발주수량 
		ggoSpread.SpreadLock 	C_SO_QTY		, -1, -1	'수주수량 
		ggoSpread.SpreadLock 	C_PO_LC_QTY		, -1, -1	'수입L/C수량 
		ggoSpread.SpreadLock 	C_SO_LC_QTY		, -1, -1	'수출L/C수량 
		ggoSpread.SpreadLock 	C_SO_REQ_QTY	        , -1, -1	'출하요청수량 
		ggoSpread.SpreadLock 	C_SO_ISSUE_QTY	        , -1, -1	'출고수량 
		ggoSpread.SpreadLock 	C_SO_CC_QTY		, -1, -1	'수출통관수량 
		ggoSpread.SpreadLock 	C_PO_CC_QTY		, -1, -1	'수입통관수량 
		ggoSpread.SpreadLock 	C_PO_RCPT_QTY	        , -1, -1	'입고수량 
		ggoSpread.SpreadLock 	C_SO_BILL_QTY	        , -1, -1	'매출수량 
		ggoSpread.SpreadLock 	C_PO_IV_QTY		, -1, -1	'매입수량 
		ggoSpread.SpreadLock 	C_PO_NO		        , -1, -1	'고객주문번호 
		ggoSpread.SpreadLock 	C_PO_SEQ_NO		, -1, -1	'순번 
		ggoSpread.SpreadLock 	C_BP_ITEM_CD	        , -1, -1	'고객품목 
		ggoSpread.SpreadLock 	C_BP_ITEM_NM	        , -1, -1	'고객품목명 


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

		ggoSpread.SSSetProtected 	C_PO_COMPANY	        , pvStartRow, pvEndRow		'발주법인 
		ggoSpread.SSSetProtected 	C_PO_COMPANY_NM	        , pvStartRow, pvEndRow		'발주법인명 
		ggoSpread.SSSetProtected 	C_SO_NO		        , pvStartRow, pvEndRow		'수주번호 
		ggoSpread.SSSetProtected 	C_SO_SEQ_NO		, pvStartRow, pvEndRow		'수주순번 
		ggoSpread.SSSetProtected 	C_ITEM_CD		, pvStartRow, pvEndRow		'품목 
		ggoSpread.SSSetProtected 	C_ITEM_NM		, pvStartRow, pvEndRow		'품목명 
		ggoSpread.SSSetProtected 	C_SPEC		        , pvStartRow, pvEndRow		'품목규격 
		ggoSpread.SSSetProtected 	C_PO_STS		, pvStartRow, pvEndRow		'발주법인상태 
		ggoSpread.SSSetProtected 	C_SO_STS		, pvStartRow, pvEndRow		'수주법인상태 
		ggoSpread.SSSetProtected 	C_UNIT		        , pvStartRow, pvEndRow		'단위 
		ggoSpread.SSSetProtected 	C_PO_QTY		, pvStartRow, pvEndRow		'발주수량 
		ggoSpread.SSSetProtected 	C_SO_QTY		, pvStartRow, pvEndRow		'수주수량 
		ggoSpread.SSSetProtected 	C_PO_LC_QTY		, pvStartRow, pvEndRow		'수입L/C수량 
		ggoSpread.SSSetProtected 	C_SO_LC_QTY		, pvStartRow, pvEndRow		'수출L/C수량 
		ggoSpread.SSSetProtected 	C_SO_REQ_QTY	        , pvStartRow, pvEndRow		'출하요청수량 
		ggoSpread.SSSetProtected 	C_SO_ISSUE_QTY	        , pvStartRow, pvEndRow		'출고수량 
		ggoSpread.SSSetProtected 	C_SO_CC_QTY		, pvStartRow, pvEndRow		'수출통관수량 
		ggoSpread.SSSetProtected 	C_PO_CC_QTY		, pvStartRow, pvEndRow		'수입통관수량 
		ggoSpread.SSSetProtected 	C_PO_RCPT_QTY	        , pvStartRow, pvEndRow		'입고수량 
		ggoSpread.SSSetProtected 	C_SO_BILL_QTY	        , pvStartRow, pvEndRow		'매출수량 
		ggoSpread.SSSetProtected 	C_PO_IV_QTY		, pvStartRow, pvEndRow		'매입수량 
		ggoSpread.SSSetProtected 	C_PO_NO		        , pvStartRow, pvEndRow		'고객주문번호 
		ggoSpread.SSSetProtected 	C_PO_SEQ_NO		, pvStartRow, pvEndRow		'순번 
		ggoSpread.SSSetProtected 	C_BP_ITEM_CD	        , pvStartRow, pvEndRow		'고객품목 
		ggoSpread.SSSetProtected 	C_BP_ITEM_NM	        , pvStartRow, pvEndRow		'고객품목명 

		.ReDraw = True
	End With
End Sub



'============================= 2.2.3 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_PO_COMPANY	        = 1	'발주법인 
	C_PO_COMPANY_NM	        = 2	'발주법인명 
	C_SO_NO		        = 3	'수주번호 
	C_SO_SEQ_NO		= 4	'수주순번 
	C_ITEM_CD		= 5	'품목 
	C_ITEM_NM		= 6	'품목명 
	C_SPEC		        = 7	'품목규격 
	C_PO_STS		= 8	'발주법인상태 
	C_SO_STS		= 9	'수주법인상태 
	C_UNIT		        = 10	'단위 
	C_PO_QTY		= 11	'발주수량 
	C_SO_QTY		= 12	'수주수량 
	C_PO_LC_QTY		= 13	'수입L/C수량 
	C_SO_LC_QTY		= 14	'수출L/C수량 
	C_SO_REQ_QTY	        = 15	'출하요청수량 
	C_SO_ISSUE_QTY	        = 16	'출고수량 
	C_SO_CC_QTY		= 17	'수출통관수량 
	C_PO_CC_QTY		= 18	'수입통관수량 
	C_PO_RCPT_QTY	        = 19	'입고수량 
	C_SO_BILL_QTY	        = 20	'매출수량 
	C_PO_IV_QTY		= 21	'매입수량 
	C_PO_NO		        = 22	'고객주문번호 
	C_PO_SEQ_NO		= 23	'순번 
	C_BP_ITEM_CD	        = 24	'고객품목 
	C_BP_ITEM_NM	        = 25	'고객품목명 
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

			C_PO_COMPANY	        = iCurColumnPos(1)	'발주법인 
			C_PO_COMPANY_NM	        = iCurColumnPos(2)	'발주법인명 
			C_SO_NO		        = iCurColumnPos(3)	'수주번호 
			C_SO_SEQ_NO		= iCurColumnPos(4)	'수주순번 
			C_ITEM_CD		= iCurColumnPos(5)	'품목 
			C_ITEM_NM		= iCurColumnPos(6)	'품목명 
			C_SPEC		        = iCurColumnPos(7)	'품목규격 
			C_PO_STS		= iCurColumnPos(8)	'발주법인상태 
			C_SO_STS		= iCurColumnPos(9)	'수주법인상태 
			C_UNIT		        = iCurColumnPos(10)	'단위 
			C_PO_QTY		= iCurColumnPos(11)	'발주수량 
			C_SO_QTY		= iCurColumnPos(12)	'수주수량 
			C_PO_LC_QTY		= iCurColumnPos(13)	'수입L/C수량 
			C_SO_LC_QTY		= iCurColumnPos(14)	'수출L/C수량 
			C_SO_REQ_QTY	        = iCurColumnPos(15)	'출하요청수량 
			C_SO_ISSUE_QTY	        = iCurColumnPos(16)	'출고수량 
			C_SO_CC_QTY		= iCurColumnPos(17)	'수출통관수량 
			C_PO_CC_QTY		= iCurColumnPos(18)	'수입통관수량 
			C_PO_RCPT_QTY	        = iCurColumnPos(19)	'입고수량 
			C_SO_BILL_QTY	        = iCurColumnPos(20)	'매출수량 
			C_PO_IV_QTY		= iCurColumnPos(21)	'매입수량 
			C_PO_NO		        = iCurColumnPos(22)	'고객주문번호 
			C_PO_SEQ_NO		= iCurColumnPos(23)	'순번 
			C_BP_ITEM_CD	        = iCurColumnPos(24)	'고객품목 
			C_BP_ITEM_NM	        = iCurColumnPos(25)	'고객품목명 

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
' Function Name : ChangeCheck
' Function Desc :
'=======================================================================================================
Function ChangeCheck()
	ChangeCheck = False

	Dim i
	Dim strInsertMark
	Dim strDeleteMark
	Dim strUpdateMark

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or ChangeCheck = True Then
        ChangeCheck = True
    End If
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

	Call InitVariables

	Call SetDefaultVal

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
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
 Sub txtSo_Frdt_DblClick(Button)
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
 Sub txtSo_Todt_DblClick(Button)
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


	'-----------------------
	'Check condition area
	'-----------------------
	If Not chkField(Document, "1") Then											'This function check indispensable field
		Exit Function
	End If



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
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
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

			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.value)        '수주법인코드 
			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '발주법인코드 
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'수주일 From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'수주일 To

			if .rdoPostFlag2(0).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & ""
			elseif .rdoPostFlag2(1).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "N"
			else
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			End if

			strVal = strVal & "&lgPageNo=" & lgPageNo                  		'☜: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	Else
	        	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.value)        '수주법인코드 
			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '발주법인코드 
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'수주일 From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'수주일 To

			if .rdoPostFlag2(0).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & ""
			elseif .rdoPostFlag2(1).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "N"
			else
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			End if

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
	Call SetToolBar("11000000000011")				'버튼 툴바 제어 

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

		    lgIntFlgModeM = Parent.OPMD_UMODE
		    lgIntFlgMode = Parent.OPMD_UMODE
		End If
	End With
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtSupplierCd.focus
	End If
	Set gActiveElement = document.activeElement
    DbQueryOk = true
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
			Case ggoSpread.UpdateFlag			'☜: 수정, 신규 
				strVal = strVal & lRow & parent.gColSep	'☜: U=Update

				'상단 스프레드 
				.vspdData.Col =C_CFM_YN		 	'확정여부 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col =C_PO_COMPANY    		'발주법인 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col =C_SO_COMPANY       	'수주법인 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col =C_PO_NO              	'발주번호(고객주문번호)
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep


				strVal = strVal & parent.gRowSep

		End Select

		lGrpCnt = lGrpCnt + 1
	Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

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


'------------------------------------------  OpenSupplier()  -------------------------------------------------


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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>멀티컴퍼니 수발주진행조회</font></td>
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
								<TD CLASS="TD5" NOWRAP>수주법인</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtSpplCd"  SIZE=10 MAXLENGTH=10 ALT="수주법인"  tag="14X">
										       <INPUT TYPE=TEXT Name="txtSpplNm" SIZE=20 MAXLENGTH=18 ALT="수주법인"  tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>발주법인</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtSupplierCd"  SIZE=10 MAXLENGTH=10 ALT="발주법인"  tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
										       <INPUT TYPE=TEXT Name="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="발주법인"  tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>수주일</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/ksm112qa1_fpDateTime2_txtSo_Frdt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/ksm112qa1_fpDateTime2_txtSo_Todt.js'></script>
											</td>
										<tr>
									</table>
								</TD>
								<TD CLASS="TD5" NOWRAP>내외자처리구분</TD>
								<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS = "RADIO" name="rdoPostFlag2" id="rdoPostFlag" value="" tag = "11" checked>
											<label for="rdoPostFlag">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoPostFlag2" id="rdoPostFlagN" value="N" tag = "11" >
											<label for="rdoPostFlagN">내자</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostFlag2" id="rdoPostFlagY" value="Y" tag = "11" >
											<label for="rdoPostFlagY">외자</label>
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
				 <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
				  <TABLE <%=LR_SPACE_TYPE_60%>>
					<TR>
						<TD WIDTH=100% COLSPAN=4>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/ksm112qa1_A_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				  </TABLE>
				 </TD>
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
