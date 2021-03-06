<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : MM111QA1
'*  4. Program Name         : 멀티컴퍼니매입조회-멀티 
'*  5. Program Desc         : 멀티컴퍼니매입조회-멀티 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2005/05/23
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Kang Su Hwan
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim interface_Production


Const BIZ_PGM_ID = "KMM111QB1.asp"
Const BIZ_PGM_ID2 = "KMM111QB101.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL = "M5111MA1"
											'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
'상단 스프레드 

Dim	C_BpCd						'수주법인 
Dim C_BpNm                      '수주법인명 
Dim C_SpplIvNo                  '공급처세금계산서번호 
Dim C_IvNo                      '매입번호 
Dim C_PostedFlg                 '매입확정여부 
Dim C_IvDt                      '매입일 
Dim C_PurGrp                    '매입그룹 
Dim C_PurGrpNm                  '매입그룹명 
Dim C_IvCur                     '화폐 
Dim C_NetDocAmt                 '매입순금액 
Dim C_TotVatDocAmt              '부가세금액 
Dim C_GrossDocAmt               '매입총금액 
Dim C_IvVatType                   '부가세유형 
Dim C_IvVatTypeNm                 '부가세유형명 
Dim C_VatRt                     '부가세율 
Dim C_PayMeth                   '결제방법 
Dim C_PayMethNm                 '결제방법명 
Dim C_PaymentTerm               '가격조건 
Dim C_PaymentTermNm		        '가격조건명 
Dim C_GlType
Dim C_GlNo
Dim C_glref_pop


'하단 스프레드 
Dim C_DIvNo									'매입번호 
Dim C_IvSeqNo								'매입 일련번호 

Dim C_CustItemCd							'품목 
Dim C_CustItemNm							'품목명 
Dim C_CustItemSpec							'품목규격 
Dim C_BillQty								'수량 
Dim C_BillUnit								'단위 
Dim C_BillPrc								'단가 
Dim C_BillDocAmt							'금액 
Dim C_VatDocAmt								'부가세금액 
Dim C_VatType								'부가세유형 
Dim C_VatTypeNm								'부가세유형명 
Dim C_VatRate								'부가세율 
Dim C_VatInc								'부가세포함구분 

Dim C_ParentRowNo							'상위 row 번호 
Dim C_Flag									'자기 번호 



Dim lgSpdHdrClicked	'2003-03-01 Release 추가 
Dim lblnWinEvent
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
Dim lgCurrRow
Dim strInspClass

Dim lgPageNo1

Dim EndDate, StartDate,CurrDate, iDBSYSDate

' === 2005.07.22 수정 ===========================================================
StartDate   = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
StartDate   = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate     = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
' === 2005.07.22 수정 ===========================================================


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

    '###검사분류별 변경부분 Start###
    strInspClass = "R"
	'###검사분류별 변경부분 End###
    'ggoSpread.ClearSpreadData
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

	frm1.txtBillFrDt.Text=StartDate
	frm1.txtBillToDt.Text=EndDate
	frm1.txtIvFrDt.Text=StartDate
	frm1.txtIvToDt.Text=EndDate

	Call SetToolbar("1100000000001111")

    frm1.txtSoCompanyCd.focus

    Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()

	With frm1.vspdData

	ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030901",,Parent.gAllowDragDropSpread

	.ReDraw = false

    .MaxCols = C_glref_pop + 1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit		C_BpCd									, "수주법인"	, 18
		ggoSpread.SSSetEdit		C_BpNm          						, "수주법인명"	, 18
		ggoSpread.SSSetEdit		C_SpplIvNo      						, "공급처세금계산서번호"	, 18
		ggoSpread.SSSetEdit		C_IvNo          						, "매입번호"	, 18
		ggoSpread.SSSetEdit		C_PostedFlg     						, "매입확정여부"	, 18
		ggoSpread.SSSetDate		C_IvDt          						, "매입일"	, 18,		2,					parent.gDateFormat
		ggoSpread.SSSetEdit		C_PurGrp        						, "매입그룹"	, 18
		ggoSpread.SSSetEdit		C_PurGrpNm      						, "매입그룹명"	, 18
		ggoSpread.SSSetEdit		C_IvCur         						, "화폐"	, 18
		ggoSpread.SSSetFloat		C_NetDocAmt     						, "매입순금액"	, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_TotVatDocAmt  						, "부가세금액"	, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_GrossDocAmt   						, "매입총금액"	, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_IvVatType       						, "부가세유형"	, 18
		ggoSpread.SSSetEdit		C_IvVatTypeNm     						, "부가세유형명"	, 18
		ggoSpread.SSSetFloat		C_VatRt         						, "부가세율"	, 12,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_PayMeth       						, "결제방법"	, 18
		ggoSpread.SSSetEdit		C_PayMethNm     						, "결제방법명"	, 18
		ggoSpread.SSSetEdit		C_PaymentTerm   						, "가격조건"	, 18
		ggoSpread.SSSetEdit		C_PaymentTermNm							, "가격조건명"	, 18
		ggoSpread.SSSetEdit 	C_GlType	, "C_GlType", 10
		ggoSpread.SSSetEdit 	C_GlNo		, "전표번호",20
		ggoSpread.SSSetButton 	C_glref_pop

		Call ggoSpread.MakePairsColumn(C_GlNo,C_glref_pop)
		Call ggoSpread.SSSetColHidden(C_GlType,C_GlType,True)
		Call ggoSpread.SSSetColHidden(C_glref_pop,C_glref_pop,True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)


    Call SetSpreadLock
    .ReDraw = true

    End With
End Sub

Sub InitSpreadSheet2()
	Call InitSpreadPosVariables2()
    With frm1

		.vspdData2.ReDraw = false

		ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread

	   .vspdData2.MaxCols = C_Flag+1
	   .vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("B")


		ggoSpread.SSSetEdit 	C_DIvNo						, "BL번호"				, 15
		ggoSpread.SSSetEdit 	C_IvSeqNo					, "순번"				, 8

		ggoSpread.SSSetEdit 	C_CustItemCd				, "품목"				, 15
		ggoSpread.SSSetEdit	    C_CustItemNm				, "품목명"				, 20
		ggoSpread.SSSetEdit		C_CustItemSpec				, "품목규격"			, 12
		ggoSpread.SSSetFloat		C_BillQty					, "수량"				, 12,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_BillUnit					, "단위"				, 12
		ggoSpread.SSSetFloat		C_BillPrc					, "단가"				, 15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_BillDocAmt				, "금액"				, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_VatDocAmt					, "부가세금액"			, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_VatType					, "부가세유형"			, 12
		ggoSpread.SSSetEdit		C_VatTypeNm					, "부가세유형명"		, 20
		ggoSpread.SSSetFloat		C_VatRate					, "부가세율"			, 12,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_VatInc					, "부가세포함구분"		, 15

		ggoSpread.SSSetEdit		C_ParentRowNo				, "C_ParentRowNo"		, 5
		ggoSpread.SSSetEdit		C_Flag						, "C_Flag"				, 5

		Call ggoSpread.SSSetColHidden(C_ParentRowNo,C_ParentRowNo, True)
 		Call ggoSpread.SSSetColHidden(C_Flag, C_Flag+1, True)

		.vspdData2.ReDraw = True

    End With
	Call SetSpreadLock2()
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
 Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData

	    ggoSpread.SpreadLock		C_BpCd         	,	-1,	C_BpCd         	-1
	    ggoSpread.SpreadLock		C_BpNm         	,	-1,	C_BpNm         	-1
		ggoSpread.SpreadLock 		C_SpplIvNo     	,	-1,	C_SpplIvNo     	-1
		ggoSpread.SpreadLock 		C_IvNo         	,	-1,	C_IvNo         	-1
		ggoSpread.SpreadLock 		C_PostedFlg    	,	-1, C_PostedFlg    	-1
		ggoSpread.SpreadLock 		C_IvDt         	,	-1, C_IvDt         	-1
		ggoSpread.SpreadLock		C_PurGrp       	,	-1, C_PurGrp       	-1
		ggoSpread.SpreadLock		C_PurGrpNm     	,	-1, C_PurGrpNm     	-1
		ggoSpread.SpreadLock 		C_IvCur        	,	-1, C_IvCur        	-1
		ggoSpread.SpreadLock		C_NetDocAmt    	,	-1, C_NetDocAmt    	-1
		ggoSpread.SpreadLock 		C_TotVatDocAmt 	,	-1, C_TotVatDocAmt 	-1
		ggoSpread.SpreadLock 		C_GrossDocAmt  	,	-1, C_GrossDocAmt  	-1
		ggoSpread.SpreadLock 		C_IvVatType    	,	-1, C_IvVatType    	-1
		ggoSpread.SpreadLock 		C_IvVatTypeNm  	,	-1, C_IvVatTypeNm  	-1
		ggoSpread.SpreadLock        C_VatRt        	,   -1, C_VatRt         -1
		ggoSpread.SpreadLock        C_PayMeth      	,   -1, C_PayMeth       -1
		ggoSpread.SpreadLock        C_PayMethNm    	,   -1, C_PayMethNm     -1
		ggoSpread.SpreadLock        C_PaymentTerm  	,   -1, C_PaymentTerm   -1
		ggoSpread.SpreadLock        C_PaymentTermNm	,   -1, C_PaymentTermNm	-1
		ggoSpread.SpreadLock 		C_GlType		,	-1, C_GlType,		-1
		ggoSpread.SpreadLock 		C_GlNo			,	-1, C_GlNo,			-1
'		ggoSpread.SpreadLock 		C_glref_pop 	,	-1,	C_glref_pop,	-1
		ggoSpread.SSSetProtected	C_glref_pop + 1	,  -1

    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadLock2()
    With frm1

    .vspdData2.ReDraw = False

    ggoSpread.Source = frm1.vspdData2
    ggoSpread.SpreadLock 1 , -1

	.vspdData2.ReDraw = True
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_BpCd				= 1		   '수주법인 
	C_BpNm              = 2        '수주법인명 
	C_SpplIvNo          = 3        '공급처세금계산서번호 
	C_IvNo              = 4        '매입번호 
	C_PostedFlg         = 5        '매입확정여부 
	C_IvDt              = 6        '매입일 
	C_PurGrp            = 7        '매입그룹 
	C_PurGrpNm          = 8        '매입그룹명 
	C_IvCur             = 9        '화폐 
	C_NetDocAmt         = 10       '매입순금액 
	C_TotVatDocAmt      = 11       '부가세금액 
	C_GrossDocAmt       = 12       '매입총금액 
	C_IvVatType         = 13       '부가세유형 
	C_IvVatTypeNm       = 14       '부가세유형명 
	C_VatRt             = 15       '부가세율 
	C_PayMeth           = 16       '결제방법 
	C_PayMethNm         = 17       '결제방법명 
	C_PaymentTerm       = 18       '가격조건 
	C_PaymentTermNm		= 19       '가격조건명 
	C_GlType    = 20     '전표 type
	C_GlNo		= 21     '전표번호 
	C_glref_pop = 22     '전표조회 팝업 

End Sub

Sub InitSpreadPosVariables2()
	C_DIvNo			= 1
	C_IvSeqNo		= 2

	C_CustItemCd	= 3         '품목 
	C_CustItemNm	= 4         '품목명 
	C_CustItemSpec	= 5         '품목규격 
	C_BillQty		= 6         '수량 
	C_BillUnit		= 7         '단위 
	C_BillPrc		= 8         '단가 
	C_BillDocAmt	= 9         '금액 
	C_VatDocAmt		= 10        '부가세금액 
	C_VatType		= 11        '부가세유형 
	C_VatTypeNm		= 12        '부가세유형명 
	C_VatRate		= 13	    '부가세율 
	C_VatInc		= 14        '부가세포함구분 

	C_ParentRowNo   = 15        '상위 row 번호 
	C_Flag          = 16        '자기 번호 
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
				C_BpCd				= 		iCurColumnPos(1)		'수주법인 
				C_BpNm              = 		iCurColumnPos(2)        '수주법인명 
				C_SpplIvNo          = 		iCurColumnPos(3)        '공급처세금계산서번호 
				C_IvNo              = 		iCurColumnPos(4)        '매입번호 
				C_PostedFlg         = 		iCurColumnPos(5)        '매입확정여부 
				C_IvDt              = 		iCurColumnPos(6)        '매입일 
				C_PurGrp            = 		iCurColumnPos(7)        '매입그룹 
				C_PurGrpNm          = 		iCurColumnPos(8)        '매입그룹명 
				C_IvCur             = 		iCurColumnPos(9)        '화폐 
				C_NetDocAmt         = 		iCurColumnPos(10)       '매입순금액 
				C_TotVatDocAmt      = 		iCurColumnPos(11)       '부가세금액 
				C_GrossDocAmt       = 		iCurColumnPos(12)       '매입총금액 
				C_IvVatType         = 		iCurColumnPos(13)       '부가세유형 
				C_IvVatTypeNm       = 		iCurColumnPos(14)       '부가세유형명 
				C_VatRt             = 		iCurColumnPos(15)       '부가세율 
				C_PayMeth           = 		iCurColumnPos(16)       '결제방법 
				C_PayMethNm         = 		iCurColumnPos(17)       '결제방법명 
				C_PaymentTerm       = 		iCurColumnPos(18)       '가격조건 
				C_PaymentTermNm		= 		iCurColumnPos(19)       '가격조건명 
				C_GlType    = iCurColumnPos(20)
				C_GlNo		= iCurColumnPos(21)
				C_glref_pop = iCurColumnPos(22)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_DIvNo			=	iCurColumnPos(1)
				C_IvSeqNo		=	iCurColumnPos(2)
				C_CustItemCd	=	iCurColumnPos(3)        '품목 
				C_CustItemNm	=	iCurColumnPos(4)        '품목명 
				C_CustItemSpec	=	iCurColumnPos(5)        '품목규격 
				C_BillQty		=	iCurColumnPos(6)        '수량 
				C_BillUnit		=	iCurColumnPos(7)        '단위 
				C_BillPrc		=	iCurColumnPos(8)        '단가 
				C_BillDocAmt	=	iCurColumnPos(9)        '금액 
				C_VatDocAmt		=	iCurColumnPos(10)        '부가세금액 
				C_VatType		=	iCurColumnPos(11)        '부가세유형 
				C_VatTypeNm		=   iCurColumnPos(12)       '부가세유형명 
				C_VatRate		=	iCurColumnPos(13)	    '부가세율 
				C_VatInc		=	iCurColumnPos(14)       '부가세포함구분 

				C_ParentRowNo   =	iCurColumnPos(15)       '상위 row 번호 
				C_Flag          =	iCurColumnPos(16)       '자기 번호 
	End Select
End Sub


'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------

Function OpenPoNo()

		Dim strRet
		Dim arrParam(2)
		Dim iCalledAspName
		Dim IntRetCD

		If IsOpenPop = True Or UCase(frm1.txtCustPoNo.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

		IsOpenPop = True

		arrParam(0) = "N"  'Return Flag
		arrParam(1) = "N"  'Release Flag
		arrParam(2) = ""  'STO Flag

'		strRet = window.showModalDialog("m3111pa1.asp", Array(window.parent,arrParam), _
'				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		iCalledAspName = AskPRAspName("M3111PA1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
			IsOpenPop = False
			Exit Function
		End If

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


		IsOpenPop = False

		If strRet(0) = "" Then
			Exit Function
		Else
			Call SetPoNo(strRet(0))

		End If

End Function

Function SetPoNo(strRet)
	frm1.txtCustPoNo.value = strRet
	frm1.txtCustPoNo.Focus
End Function

'------------------------------------------  OpenSoCompany()  -------------------------------------------------
' Name : OpenSoCompany()
' Description : SpreadItem PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenSoCompany()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "수주법인"
	arrParam(1) = "B_BIZ_PARTNER"

	arrParam(2) = Trim(frm1.txtSoCompanyCd.Value)

	arrParam(4) = "BP_TYPE In ('S','CS') And IN_OUT_FLAG = 'O'"
	arrParam(5) = "수주법인"

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = "수주법인"
	arrHeader(1) = "수주법인명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		frm1.txtSoCompanyCd.focus
		Exit Function
	Else
		frm1.txtSoCompanyCd.Value= arrRet(0)
		frm1.txtSoCompanyNm.Value= arrRet(1)
		frm1.txtSoCompanyCd.focus
	End If
End Function

'------------------------------------------  OpenGLRef()  -------------------------------------------------
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName

	If lblnWinEvent = True Then Exit Function

	lblnWinEvent = True

	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	frm1.vspdData.Col = C_GlNo                      '전표번호 
    arrParam(0) = Trim(frm1.vspdData.Text)
    frm1.vspdData.Col = C_IvNo                      '매입번호 
	arrParam(1) = Trim(frm1.vspdData.Text)

   frm1.vspdData.Col = C_GlType                      '전표번호 type

   If Trim(frm1.vspdData.Text) = "A" Then               '회계전표팝업 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif Trim(frm1.vspdData.Text) = "T" Then          '결의전표팝업 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif Trim(frm1.vspdData.Text) = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '아직 전표가 생성되지 않았습니다.
    End if

	lblnWinEvent = False

End Function

'======================================   Getglno()  =====================================
Sub Getglno()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
    Dim strwhere,strrefno
    Dim strglno

    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    frm1.vspdData.Col = C_IvNo           '매입번호 
    strrefno = Trim(frm1.vspdData.Text)
    Err.Clear

    strwhere = " ref_no = '" & strrefno & "'"
    Call CommonQueryRs(" gl_no ", " a_gl ",strwhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if Trim(lgF0) = "" then
        Err.Clear
        Call CommonQueryRs(" temp_gl_no ", " a_temp_gl ",strwhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

        if Trim(lgF0) = "" then
            frm1.vspdData.Col = C_GlType
            frm1.vspdData.Text = "B"
        else
            frm1.vspdData.Col = C_GlType
            frm1.vspdData.Text = "T"
        end if

    else
        frm1.vspdData.Col = C_GlType
        frm1.vspdData.Text = "A"
    end if

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

'====================================== sprRedComColor() ======================================
'	Name : sprRedComColor()
'	Description : 발주일자가 현재 일자보다 적을떄 적색 신호...
'==============================================================================================
Sub sprRedComColor(ByVal Col, ByVal Row, ByVal Row2)
    With frm1
		.vspdData2.Col = Col
		.vspdData2.Col2 = Col
		.vspdData2.Row = Row
		.vspdData2.Row2 = Row2
		.vspdData2.ForeColor = vbRed
    End With
End Sub
'====================================== sprBlackComColor() ======================================
'	Name : sprBlackComColor()
'	Description : 발주일자가 현재 일자보다 적을떄 적색 신호...
'==============================================================================================
Sub sprBlackComColor(ByVal Col, ByVal Row, ByVal Row2)
    With frm1
		.vspdData2.Col = Col
		.vspdData2.Row = Row
        .vspdData2.ForeColor = &H0&
    End With
End Sub


'==========================================   SpplChange()  ======================================
'	Name : SpplChange()
'	Description :
'=================================================================================================

Sub SpplChange()
    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Sub
	End If

    Dim strVal
    Dim strssText1, strssText2
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim iparentrow
    Dim index
    Dim iRow

	with frm1.vspdData2
	    iRow        = .ActiveRow
		.Row		= .ActiveRow
		.Col		= C_ParentPrNo
		strssText1	= Trim(.text)
		.Col		= C_SpplCd
		strssText2	= Trim(.text)
		.Col        = C_ParentRowNo
		iparentrow  = Trim(.text)
		if strssText2 = "" then
			Exit Sub
		End If

	End with

	lngRangeFrom = DataFirstRow(iparentrow)
	lngRangeTo   = DataLastRow(iparentrow)

	for index = lngRangeFrom to lngRangeTo
	    if index <> iRow and strssText2 <> "" then
	        frm1.vspdData2.Row = index
	        frm1.vspdData2.Col = C_SpplCd
	        if UCase(strssText2) = UCase(Trim(frm1.vspdData2.text)) then
                Call DisplayMsgBox("17A005","X" ,"공급처", "X")
                frm1.vspdData2.Row = iRow
	            frm1.vspdData2.Col = C_SpplCd
	            frm1.vspdData2.text = ""
 	            Exit sub
	        End If
	    End If
	next

    strVal = BIZ_PGM_ID & "?txtMode=" & "LookSppl"
    strVal = strVal & "&txtPrNo=" & strssText1
    strVal = strVal & "&txtBpCd=" & strssText2

    If LayerShowHide(1) = False Then Exit Sub

	Call RunMyBizASP(MyBizASP, strVal)
End Sub

'=======================================================================================================
'   Sub Name : SheetFocus
'   Sub Desc :
'=======================================================================================================
Sub SheetFocus(Byval iChildRow)
	Dim iParentRow
	Dim CheckField1
	Dim CheckField2
	Dim i
	Dim lngStart
	Dim lngEnd
	Dim strSampleNo
	Dim strFlag

	With frm1.vspdData2
		.Row = iChildRow
		.Col = C_ParentRowNo
		iParentRow = CLng(.Text)
		.Col = C_SpplCd
		strSampleNo = .Text
		.Col = C_Flag
		strFlag = .Text
	End With

	Call ParentGetFocusCell(iParentRow, strSampleNo, strFlag)
End Sub

'=======================================================================================================
'   Event Name : ParentGetFocusCell
'   Event Desc :
'=======================================================================================================
Sub ParentGetFocusCell(ByVal ParentRow, ByVal strSampleNo, Byval strFlag)
	Dim CheckField1
	Dim CheckField2
	Dim i
	Dim lngStart
	Dim lngEnd

	With frm1.vspdData
		.Row = ParentRow
		.Col = 1
		.Action = 0		'Active Cell
	End With

	With frm1.vspdData2
		.ReDraw = False
		lngStart = ShowFromData(ParentRow, lglngHiddenRows(ParentRow - 1))
		.ReDraw = True
		lngEnd = lngStart + lglngHiddenRows(ParentRow - 1) - 1
		For i = lngStart To lngEnd
			.Row = i
			.Col = C_SpplCd
			CheckField1 = .Text
			.Col = C_Flag
			CheckField2 = .Text
			If CheckField1 = strSampleNo And CheckField2 = strFlag Then
				Exit For
			End If
		Next

	End With

	Set gActiveElement = document.activeElement

End Sub

'=======================================================================================================
'   Function Name : ShowFromData
'   Function Desc :
'=======================================================================================================
Function ShowFromData(Byval Row, Byval lngShowingRows)	'###그리드 컨버전 주의부분###
'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 3을 리턴하는 기능을 수행하는 함수다.
	ShowFromData = 0
	Dim lngRow
	Dim lngStartRow

	With frm1.vspdData2

		Call SortSheet()
		'------------------------------------
		' Find First Row
		'------------------------------------
		lngStartRow = 0
'check this !
		If .MaxRows < 1 Then Exit Function

		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo
			If Row = CInt(.Text) Then
				lngStartRow = lngRow
				ShowFromData = lngRow
				Exit For
			End If
		Next

		'------------------------------------
		' Show Data
		'------------------------------------
		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = 1
			.Row2 = .MaxRows
			.Col = C_Flag
			.Col2 = C_Flag
			.DestCol = 0
			.DestRow = 1
			.Action = 19	'SS_ACTION_COPY_RANGE
			.RowHidden = False

			.BlockMode = False

			'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 첫번째 부터 2번째 까지의 Row를 숨긴다.
			If lngStartRow > 1 Then
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

			'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 7번째 부터 마지막 까지의 Row를 숨긴다.
			If lngStartRow < .MaxRows Then
				If lngStartRow + lngShowingRows <= .MaxRows Then
					.BlockMode = True
					.Row = lngStartRow + lngShowingRows
					.Row2 = .MaxRows
					.RowHidden = True
					.BlockMode = False
				End If
			End If

			.BlockMode = False

			.Row = lngStartRow	'2003-03-01 Release 추가 
			.Col = 0			'2003-03-01 Release 추가 
			.Action = 0			'2003-03-01 Release 추가 
		End If
	End With
End Function

'=======================================================================================================
'   Function Name : DeleteDataForInsertSampleRows
'   Function Desc :
'=======================================================================================================
Function DeleteDataForInsertSampleRows(ByVal Row, Byval lngShowingRows)
	DeleteDataForInsertSampleRows = False

	Dim lngRow
	Dim lngStartRow

	With frm1.vspdData2

		Call SortSheet()

		'------------------------------------
		' Find First Row
		'------------------------------------
		lngStartRow = 0
		For lngRow = 1 To .MaxRows
			.Row = lngRow
			.Col = C_ParentRowNo
			If Row = Clng(.Text) Then
				lngStartRow = lngRow
				DeleteDataForInsertSampleRows = True
				Exit For
			End If
		Next

		'------------------------------------
		' Delete Data
		'------------------------------------
		If lngStartRow > 0 Then
			.BlockMode = True
			.Row = lngStartRow
			.Row2 = lngStartRow + lngShowingRows - 1
			.Action = 5		'5 - Delete Row 	SS_ACTION_DELETE_ROW
			'********** START
			.MaxRows = .MaxRows - lngShowingRows
			'********** END
			.BlockMode = False
		End If
	End With
End Function

'======================================================================================================
' Function Name : SortSheet
' Function Desc : This function set Muti spread Flag
'=======================================================================================================
Function SortSheet()
	SortSheet = false

    With frm1.vspdData2
        .BlockMode = True
        .Col = 0
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .SortBy = 0 'SS_SORT_BY_ROW

        .SortKey(1) = C_ParentRowNo
        .SortKey(2) = C_Flag

        .SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .SortKeyOrder(2) = 0 'SS_SORT_ORDER_ASCENDING

        .Col = 1	'C_SupplierCd	'###그리드 컨버전 주의부분###
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .Action = 25 'SS_ACTION_SORT

        .BlockMode = False
    End With
    SortSheet = true
End Function

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

'=======================================================================================================
' Function Name : InsertSampleRows
' Function Desc :
'=======================================================================================================
Sub InsertSampleRows()
	Dim i
	Dim j
	Dim lngMaxRows
	Dim strInspItemCd
	Dim strInspSeries
	Dim lngOldMaxRows
	Dim strMark
	Dim lRow

    With frm1
    	If .vspdData.Row < 1 Then
    		Exit Sub
    	End If

   		Call LayerShowHide(1)

    	lRow = .vspdData.ActiveRow
    	' 해당 검사항목/차수를 가지고 있는 측정치들 삭제 
    	Call DeleteDataForInsertSampleRows(lRow, lglngHiddenRows(lRow - 1))

    	' 행 추가 
    	lngOldMaxRows = .vspdData2.MaxRows

    	.vspdData.Row = lRow
    	.vspdData.Col = C_ApportionQty
    	lngMaxRows = UNICDbl(.vspdData.Text)
    	.vspdData2.MaxRows = lngOldMaxRows + lngMaxRows

	End With

    ggoSpread.Source = frm1.vspdData2
    strMark = ggoSpread.InsertFlag

    With frm1.vspdData2
		.BlockMode = True
		.Row = lngOldMaxRows + 1
		.Row2 = .MaxRows
		.Col = C_ParentRowNo
		.Col2 = C_ParentRowNo
		.Text = lRow
		.BlockMode = False

		j = 0
        For i = lngOldMaxRows + 1 To .MaxRows
			j = j + 1
			.Row = i
			.Col = 0
			.Text = strMark
			'********** START
			.Col = C_Flag
			.Text = strMark
			'********** END
			.Col = C_SupplierCd
			.Text = j
		Next
	End With

	frm1.vspdData.Col = C_InspUnitIndctnCd

	Call SetSpreadColor2byInspUnitIndctn(lngOldMaxRows + 1, frm1.vspdData2.MaxRows, frm1.vspdData.Text, "I")

	frm1.vspdData2.Row = lngOldMaxRows + 1
	frm1.vspdData2.Col = C_SpplCd
	frm1.vspdData2.Action = 0
	lglngHiddenRows(lRow - 1) = lngMaxRows
    Call LayerShowHide(0)
End Sub

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
	IF Col = C_glref_pop then
       Call Getglno()
       Call OpenGLRef()
    End If
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

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()	'###그리드 컨버전 주의부분###

	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
'	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")  'Lock  Suitable  Field
	Call InitSpreadSheet
	Call InitSpreadSheet2
	Call InitVariables
	Call SetDefaultVal
	set gActiveSpdSheet = frm1.vspdData
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
		If DbQuery2(lgCurrRow, False) = False Then	Exit Sub
	End If
	frm1.vspdData.redraw = true
End Sub



'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================
'==========================================================================================
'   Event Name : txtBillFrDt
'   Event Desc :
'==========================================================================================
Sub txtBillFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBillFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtBillFrDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtBillToDt
'   Event Desc :
'==========================================================================================
Sub txtBillToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBillToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtBillToDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtIvFrDt
'   Event Desc :
'==========================================================================================
Sub txtIvFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIvFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtIvFrDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtBillToDt
'   Event Desc :
'==========================================================================================
Sub txtIvToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtIvToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtIvToDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtBillFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtBillToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtIvFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtIvToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


'======================================================================================================
' Function Name : FncQuery
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
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If

' === 2005.07.22 수정 ===========================================================

	If ValidDateCheck(frm1.txtBillFrDt, frm1.txtBillToDt) = False Then Exit Function
	If ValidDateCheck(frm1.txtIvFrDt, frm1.txtIvToDt) = False Then Exit Function

' === 2005.07.22 수정 ===========================================================

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

    Dim IntRetCD

	'-----------------------
    'Precheck area
    '-----------------------
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

    '8월 정기패치: 화면에 보이는 우측 스프레드에 행추가 되었으나 Hidden 스프레드에 반영이 안된 것 체크 START
	If DefaultCheck = False Then
		Exit Function
	End If
    '8월 정기패치: 화면에 보이는 우측 스프레드에 행추가 되었으나 Hidden 스프레드에 반영이 안된 것 체크 END

'	intRetCd = DisplayMsgBox("900018", VB_YES_NO, "X", "X")   '☜ 바쾪E觀?
'	If intRetCd = VBNO Then
'		Exit Function
'	End IF


    '-----------------------
    'Check content area
    '-----------------------
'    If Not chkField(Document, "1") Then
'       		Exit Function
'    End If

    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then
		Exit Function
	End If

	Set gActiveElement = document.activeElement
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
	Set gActiveElement = document.activeElement
End Sub

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
	with frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then

		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtBillFrDt=" & Trim(.txtBillFrDt.Text)
			strVal = strVal & "&txtBillToDt=" & Trim(.txtBillToDt.Text)
			strVal = strVal & "&txtIvFrDt=" & Trim(.txtIvFrDt.Text)
			strVal = strVal & "&txtIvToDt=" & Trim(.txtIvToDt.Text)
			strVal = strVal & "&rdoCfmflg=" & Trim(.rdoCfmflg.Text)
			strVal = strVal & "&txtCustPoNo=" & Trim(.txtCustPoNo.value)
			strVal = strVal & "&txtBlNo=" & Trim(.txtBlNo.value)
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '☜: Next key tag
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else

		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtBillFrDt=" & Trim(.txtBillFrDt.Text)
			strVal = strVal & "&txtBillToDt=" & Trim(.txtBillToDt.Text)
			strVal = strVal & "&txtIvFrDt=" & Trim(.txtIvFrDt.Text)
			strVal = strVal & "&txtIvToDt=" & Trim(.txtIvToDt.Text)
			if .rdoCfmflg(0).checked = true then
				strVal = strVal & "&rdoCfmflg=" & "%"
			elseif .rdoCfmflg(1).checked = true then
				strVal = strVal & "&rdoCfmflg=" & "Y"
			else
				strVal = strVal & "&rdoCfmflg=" & "N"
			End if
			strVal = strVal & "&txtCustPoNo=" & Trim(.txtCustPoNo.value)
			strVal = strVal & "&txtBlNo=" & Trim(.txtBlNo.value)
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '☜: Next key tag
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
	Dim index

	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("11001000000011")				'버튼 툴바 제어 

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
		frm1.txtSoCompanyCd.focus
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
	Dim pRow

	'/* 9월 정기패치: 좌측 스프레드의 행간 이동 시 이미 조회된 자료나 입력된 자료를 읽어 들일 때에도 '' 창 띄우기 - START */
	Call LayerShowHide(1)

	With frm1
		.vspdData.redraw = false
		.vspdData.Row = CInt(Row)
		.vspdData.Col = .vspdData.MaxCols
		pRow  = CInt(.vspdData.Text)

		If lglngHiddenRows(pRow - 1) <> 0 And NextQueryFlag = False Then
			.vspdData2.ReDraw = False
			lngRet = ShowFromData(pRow, lglngHiddenRows(pRow - 1))	'ex) 첫번째 그리드의 특정 Row에 해당하는 두번째 그리드의 Row수가 10개일때 보여줄 데이터가 3번째 부터 6번째까지 4개이면 3을 리턴하는 기능을 수행하는 함수다.
			Call SetToolBar("11001111001011")				'버튼 툴바 제어 
			Call LayerShowHide(0)
			.vspdData2.ReDraw = True
			DbQuery2 = True
			.vspdData.redraw = True
			Exit Function
		End If

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		.vspdData.Row = Row
		.vspdData.Col = C_IvNo
		strVal = strVal & "&txtIvNo=" & trim(.vspdData.text)
		strVal = strVal & "&lgStrPrevKeyM="  & lgStrPrevKeyM(Row - 1)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'☜: Next key tag
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(pRow)

		'msgbox strVal
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

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_Flag

		.Col2 = C_Flag
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False
	End With


	For Index = lngRangeFrom to lngRangeTo
    	frm1.vspdData2.Row = Index
'    	Call checkdt(Index)
    	If Index = lngRangeTo Then
				frm1.vspdData2.Row = Index
				frm1.vspdData2.Col = 1
				frm1.vspdData2.Action = 0
				frm1.vspdData2.focus
		End if
	Next

	DbQueryOk2 = true

End Function


Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave()
	DbSave = False                                                          '⊙: Processing is NG
	Dim lRow
	Dim lGrpCnt
	Dim strVal,strIU, strDel
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim parentRow
	Dim iReqQty,totalQty,totalRate
	Dim lgTransSep
	Dim lgHdDtlSep
	Dim strValUp, strReqNo, strDlvyDt, strModifyChk, iRowMode
	Dim iStrPurOrg
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]
	Dim iColSep,iRowSep
	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규]
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제]
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

	Dim intRetCd

	Dim chknum
	chknum = 0
	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002
	End With

	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 1
	strVal = ""
    strDel = ""
    strIU  = ""
    lgTransSep = "º"
    lgHdDtlSep = "Ð"
    iRowMode = ""

	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	'-----------------------
	'Data manipulate area
	'-----------------------
	With frm1
	    For parentRow = 1 To .vspdData.MaxRows
		    iStrPurOrg=""
			If Trim(GetSpreadValue(.vspdData,C_Check,parentRow,"X","X")) = 1 Then

			    lngRangeFrom = DataFirstRow(parentRow)
			    lngRangeTo   = DataLastRow(parentRow)

			    '-----상단 스프레드 값을 담는다. -------------------------
			    iRowMode = Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))

			    If iRowMode = ggoSpread.UpdateFlag Then
					strValUp = "UPDATE" & iColSep
				End If

				strValUp = strValUp & Trim(GetSpreadText(.vspdData,C_ReqNo,parentRow,"X","X")) & iColSep


				If Trim(GetSpreadText(.vspdData,C_BlDocAmt,parentRow,"X","X"))="" Then
					strValUp = strValUp & "0" & iColSep
				Else
					strValUp = strValUp & UNIConvNum(Trim(GetSpreadText(.vspdData,C_BlDocAmt,parentRow,"X","X")),0) & iColSep
				End If

				strValUp = strValUp & Trim(GetSpreadText(.vspdData,C_Unit,parentRow,"X","X")) & iColSep

				If iRowMode = ggoSpread.UpdateFlag AND _
					CDate(UNIConvDate(Trim(GetSpreadText(.vspdData,C_DlvyDt,parentRow,"X","X")))) < CDate(UNIConvDate(Trim(CurrDate))) Then
				    Call DisplayMsgBox("172120","X", parentRow & "행 ","X")
					Call LayerShowHide(0)
					Call RemovedivTextArea
'msg modify 20040506 by kjt
					.vspdData.Row = ParentRow
					.vspdData.Col = C_DlvyDt
					.vspdData.Action = 0		'Active Cell
					Call DbQuery2(ParentRow,False)
					Exit Function
				End If
'' 2004 04 13 update by kjt
				If iRowMode = ggoSpread.DeleteFlag Then
					For lRow = lngRangeFrom To lngRangeTo
						frm1.vspddata2.Row = lRow
						frm1.vspddata2.Col = 0
						if frm1.vspddata2.text = ggoSpread.InsertFlag Then
							intRetCd = DisplayMsgBox("900038", Parent.VB_YES_NO, "X", "X")
							If intRetCd = VBNO Then
								Call LayerShowHide(0)
								frm1.vspdData.Row = parentRow
								frm1.vspdData.Col = 0
								frm1.vspdData.text = ggoSpread.UpdateFlag
								frm1.vspdData.Col = 1
								frm1.vspdData.Action = 0
								frm1.vspdData.focus
								Call DbQuery2(parentRow,False)
								Exit Function
							End IF
						End if
					Next
				End If

				strValUp = strValUp & strDlvyDt & iColSep
				strValUp = strValUp & Trim(GetSpreadText(.vspdData,C_ORGCd,parentRow,"X","X")) & iColSep
				strValUp = strValUp & parentRow & lgHdDtlSep	'7 라인 
				strVal = strValUp
				'----------------------------------------------------------
			    totalQty  = 0
			    totalRate = 0

			    If lngRangeTo > 0 Then
					For lRow = lngRangeFrom To lngRangeTo
						chknum = chknum + 1
						If CheckDuplSppl(lRow) = False Then
							DbSave = False
							Call LayerShowHide(0)
							Call RemovedivTextArea
							Exit Function
						End If
						.vspddata2.row = lRow
						.vspddata2.col = C_SpplCd
					    If Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) <> ggoSpread.DeleteFlag Then
							totalQty = totalQty + Unicdbl(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))
							totalRate = totalRate + Unicdbl(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))
					    End If

					    Select Case GetSpreadText(.vspdData2,0,lRow,"X","X")

							Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
							    If GetSpreadText(.vspdData2,0,lRow,"X","X")=ggoSpread.InsertFlag then
									strIU = strIU & "C" & iColSep
								Else
									strIU = strIU & "U" & iColSep
								End If

					            strIU = strIU & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep

							    If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
									strIU = strIU & "0" & iColSep
								Else
									strIU = strIU & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
								End If

								If Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))="" Then
									strIU = strIU & "0" & iColSep
								Else
									strIU = strIU & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X")),0) & iColSep
								End If

								If CDate(UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X")))) < CDate(UNIConvDate(Trim(CurrDate))) Then
								    Call DisplayMsgBox("172140","X", strReqNo & " - " & chknum  & chr(32) & " 행 ","X")
								    Call LayerShowHide(0)
								    Call RemovedivTextArea
									' move to error row & col 2004-05-07 update by jt.kim
									.vspdData.Row = ParentRow
									.vspdData.Col = 1
									.vspdData.Action = 0		'Active Cell
									Call DbQuery2(ParentRow,False)
								    .vspdData2.Row = lRow
								    .vspdData2.Col = C_PlanDt
									.vspdData2.Action = 0		'Active Cell
								    Exit Function
								End If

								If CDate(UNIConvDate(Trim(strDlvyDt))) <  CDate(UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X")))) Then
								    Call DisplayMsgBox("172125","X", strReqNo & " - " & chknum  & chr(32) & " 행 ","X")
								    Call LayerShowHide(0)
								    Call RemovedivTextArea
									' move to error row & col 2004-05-07 update by jt.kim
									.vspdData.Row = ParentRow
									.vspdData.Col = C_DlvyDt
									.vspdData.Action = 0		'Active Cell
									Call DbQuery2(ParentRow,False)
								    .vspdData2.Row = lRow
								    .vspdData2.Col = C_PlanDt
									.vspdData2.Action = 0		'Active Cell
								    Exit Function
								End If
								strIU = strIU & UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X"))) & iColSep
								strIU = strIU & Trim(iStrPurOrg) & iColSep
								strIU = strIU & Trim("" & GetSpreadText(.vspdData2,C_GrpCd,lRow,"X","X")) & iColSep
								strIU = strIU & "" & iColSep
								strIU = strIU & parentRow & iRowSep

							Case ggoSpread.DeleteFlag				'☜: 삭제 
								strDel = strDel & "D" & iColSep			'☜: D=Delete
					            strDel = strDel & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep

							    If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
									strDel = strDel & "0" & iColSep
								Else
									strDel = strDel & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
								End If

								If Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))="" Then
									strDel = strDel & "0" & iColSep
								Else
									strDel = strDel & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X")),0) & iColSep
								End If

								strDel = strDel & UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X"))) & iColSep
								strDel = strDel & Trim("" & GetSpreadText(.vspdData2,C_GrpCd,lRow,"X","X")) & iColSep
								strDel = strDel & parentRow & iRowSep

						End Select

					Next
				Else
					totalRate=100
					totalQty=iReqQty
				End If

			   	If iRowMode = ggoSpread.UpdateFlag Then
			   		If totalRate <> 100 Then
					    Call DisplayMsgBox("171325", "X", parentRow & "행 ", "X")
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
									' move to error row & col 2004-05-07 update by jt.kim
									.vspdData.Row = ParentRow
									.vspdData.Col = 1
									.vspdData.Action = 0		'Active Cell
									Call DbQuery2(ParentRow,False)
								    .vspdData2.Row = 1
								    .vspdData2.Col = C_Quota_Rate
									.vspdData2.Action = 0		'Active Cell

					    Exit Function
					End If

			   		If totalQty <> iReqQty Then
					    Call DisplayMsgBox("172420","X",strReqNo, "X")
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
					    Exit Function
					End If

				End If

				strVal =  strVal & strDel & strIU & lgTransSep
				Select Case Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))
				    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag,ggoSpread.DeleteFlag
				         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 

				            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
				            objTEXTAREA.name = "txtCUSpread"
				            objTEXTAREA.value = Join(iTmpCUBuffer,"")
				            divTextArea.appendChild(objTEXTAREA)

				            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
				            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
				            iTmpCUBufferCount = -1
				            strCUTotalvalLen  = 0
				         End If

				         iTmpCUBufferCount = iTmpCUBufferCount + 1

				         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
				            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
				            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
				         End If
				         iTmpCUBuffer(iTmpCUBufferCount) =  strVal
				         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				End Select
			End If
			strVal  = ""
			strDel  = ""
			strIU   = ""
		Next

	End With

	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If

'	Call ExecMyBizASP(frm1, BIZ_PGM_ID)			'☜: 비지니스 ASP 를 가동 

	DbSave = True
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	frm1.vspdData2.MaxRows = 0

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If

 	with frm1
		if (UniConvDateToYYYYMMDD(.txtReqFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtReqToDt.text,Parent.gDateFormat,"")) and trim(.txtReqFrDt.text)<>"" and trim(.txtReqToDt.text)<>"" then
			Call DisplayMsgBox("17a003", "X","요청일", "X")
			Exit Function
		End If

		if (UniConvDateToYYYYMMDD(.txtDlvyFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtDlvyToDt.text,Parent.gDateFormat,"")) and trim(.txtDlvyFrDt.text)<>"" and trim(.txtDlvyToDt.text)<>"" then
			Call DisplayMsgBox("17a003", "X","필요일", "X")
			Exit Function
		End If

	End with

	'-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then
		Exit Function
	End If

End Function


'==========================================================================================
'   Event Name : btnGL_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnGL_OnClick()
       Call Getglno()
       Call OpenGLRef()
End Sub


'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	If lgIntFlgMode = parent.OPMD_CMODE Then Exit Sub

	 If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then
          Exit Sub
     End If

	Frm1.vspdData.ReDraw = False
	IF Col = C_glref_pop then
       Call Getglno()
       Call OpenGLRef()
    End If

	Frm1.vspdData.ReDraw = True
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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>멀티컴퍼니 매입조회</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
								<TD CLASS="TD5" NOWRAP>수주법인</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수주법인"   NAME="txtSoCompanyCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenSoCompany" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSoCompany()" >
														<INPUT TYPE=TEXT ALT="수주법인" NAME="txtSoCompanyNm" SIZE=20 MAXLENGTH=50 tag="24X">
								<TD CLASS="TD5" NOWRAP>공급처계산서발행일</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm111qa1_fpDateTime2_txtBillFrDt.js'></script>~
										<script language =javascript src='./js/kmm111qa1_fpDateTime2_txtBillToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>매입일</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm111qa1_fpDateTime2_txtIvFrDt.js'></script>~
										<script language =javascript src='./js/kmm111qa1_fpDateTime2_txtIvToDt.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>매입확정처리여부</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="매입확정처리여부" NAME="rdoCfmflg" id = "rdoCfmflg1" Value="A" checked tag="11"><label for="rdoCfmflg1">&nbsp;전체&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="매입확정처리여부" NAME="rdoCfmflg" id = "rdoCfmflg2" Value="Y" tag="11"><label for="rdoCfmflg2">&nbsp;확정&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="매입확정처리여부" NAME="rdoCfmflg" id = "rdoCfmflg3" Value="N" tag="11"><label for="rdoCfmflg3">&nbsp;미확정&nbsp;</label></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발주번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주번호"   NAME="txtCustPoNo" SIZE=33 MAXLENGTH=18 tag="11NXXU" >
								<!--<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()">-->
								</TD>
								<TD CLASS="TD5" NOWRAP>공급처계산서번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처계산서번호"   NAME="txtBlNo" SIZE=35 MAXLENGTH=18 tag="11NXXU" ></TD>
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
								<TD HEIGHT=50% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm111qa1_A_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=50% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm111qa1_B_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" align="left">&nbsp;
					<BUTTON NAME="btnGL" CLASS="CLSMBTN">전표조회</BUTTON>
					</TD>
					<td WIDTH="*" align="right"></td>
					<TD WIDTH=10>&nbsp;</TD>
			 </TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>