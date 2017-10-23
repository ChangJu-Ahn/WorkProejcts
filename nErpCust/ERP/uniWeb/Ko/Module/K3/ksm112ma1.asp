<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111qa1
'*  4. Program Name         : ��Ƽ���۴ϼ�����ȸ 
'*  5. Program Desc         : ��Ƽ���۴ϼ�����ȸ-��Ƽ 
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/24
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "ksm112mb1.asp"
Const BIZ_PGM_ID2 = "ksm112mb01.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL = "S3111MA1"
											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��� �������� 
Dim C_SEL_YN 	         		'���� 
Dim C_CFM_FLAG				'����Ȯ������ 
Dim C_SOLD_TO_PARTY			'���ֹ��� 
Dim C_BP_FULL_NM			'���ֹ��θ� 
Dim C_CUST_PO_NO			'�����ֹ�ȣ 
Dim C_SO_NO				'���ֹ�ȣ 
Dim C_EXPORT_FLAG			'�����ڱ��� 
Dim C_SO_DT				'������ 
Dim C_SALES_GRP				'�����׷� 
Dim C_SALES_GRP_FULL_NM			'�����׷�� 
Dim C_CUR				'ȭ�� 
Dim C_NET_AMT				'���ֱݾ� 
Dim C_VAT_AMT				'�ΰ����ݾ� 
Dim C_NET_VAT_TOTAMT			'�����ѱݾ� 
Dim C_VAT_TYPE				'�ΰ������� 
Dim C_VAT_TYPE_NM			'�ΰ��������� 
Dim C_VAT_RATE				'�ΰ����� 
Dim C_PAY_METH				'������� 
Dim C_PAY_METH_NM			'��������� 
Dim C_INCOTERMS				'�������� 
Dim C_INCOTERMS_NM			'�������Ǹ� 
Dim C_HIDDEN_CFM_FLAG			'����Ȯ������(HIDDEN)

'�ϴ� �������� 
Dim C_ITEM_CD				'ǰ�� 
Dim C_ITEM_NM				'ǰ��� 
Dim C_SPEC				'ǰ��԰� 
Dim C_CUST_ITEM_CD			'��ǰ�� 
Dim C_BP_ITEM_NM			'��ǰ��� 
Dim C_BP_ITEM_SPEC			'��ǰ��԰� 
Dim C_SO_QTY				'���� 
Dim C_SO_UNIT				'���� 
Dim C_SO_PRICE				'�ܰ� 
Dim C_NET_AMT2				'�ݾ� 
Dim C_DLVY_DT				'������ 
Dim C_VAT_AMT2				'�ΰ����ݾ� 
Dim C_VAT_RATE2				'�ΰ����� 
Dim C_VAT_TYPE2				'�ΰ������� 
Dim C_VAT_TYPE_NM2			'�ΰ��������� 
Dim C_VAT_INC_FLAG			'�ΰ������Ա��� 



Dim lgSpdHdrClicked	'2003-03-01 Release �߰� 
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim lgIntFlgModeM                 'Variable is for Operation Status

Dim lgStrPrevKeyM			'Multi���� �������� ���� ���� 
Dim lglngHiddenRows		'Multi���� �������� ���� ����	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.

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
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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

		ggoSpread.SSSetCheck    C_SEL_YN, "����",10,,,true
		ggoSpread.SSSetEdit 	C_CFM_FLAG		,"����Ȯ������",20
		ggoSpread.SSSetEdit 	C_SOLD_TO_PARTY		,"���ֹ���"		,15		'���ֹ��� 
		ggoSpread.SSSetEdit 	C_BP_FULL_NM		,"���ֹ��θ�"		,20		'���ֹ��θ� 
		ggoSpread.SSSetEdit 	C_CUST_PO_NO		,"�����ֹ�ȣ"	,15		'�����ֹ�ȣ 
		ggoSpread.SSSetEdit 	C_SO_NO			,"���ֹ�ȣ"		,15		'���ֹ�ȣ 
		ggoSpread.SSSetEdit 	C_EXPORT_FLAG		,"�����ڱ���"		,15		'�����ڱ��� 
		ggoSpread.SSSetDate 	C_SO_DT			,"������"		,		10,		2,					parent.gDateFormat'������ 
		ggoSpread.SSSetEdit 	C_SALES_GRP		,"�����׷�"		,20		'�����׷� 
		ggoSpread.SSSetEdit 	C_SALES_GRP_FULL_NM	,"�����׷��"		,20		'�����׷�� 
		ggoSpread.SSSetEdit 	C_CUR			,"ȭ��"			,20		'ȭ�� 
		ggoSpread.SSSetFloat 	C_NET_AMT		,"���ֱݾ�"			,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_VAT_AMT		,"�ΰ����ݾ�"		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_NET_VAT_TOTAMT	,"�����ѱݾ�"		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_VAT_TYPE		,"�ΰ�������"		,		10,		,					,	  5,	  2
		ggoSpread.SSSetEdit 	C_VAT_TYPE_NM		,"�ΰ���������"		,20		'�ΰ��������� 
		ggoSpread.SSSetFloat 	C_VAT_RATE		,"�ΰ�����"		,		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_PAY_METH		,"�������"		,20		'������� 
		ggoSpread.SSSetEdit 	C_PAY_METH_NM		,"���������"		,20		'��������� 
		ggoSpread.SSSetEdit 	C_INCOTERMS		,"��������"		,20		'�������� 
		ggoSpread.SSSetEdit 	C_INCOTERMS_NM		,"�������Ǹ�"		,20		'�������Ǹ� 
		ggoSpread.SSSetEdit 	C_HIDDEN_CFM_FLAG	,"����Ȯ������"		,10		'����Ȯ������(HIDDEN)


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

		ggoSpread.SSSetEdit 	C_ITEM_CD	,"ǰ��" 		,			18,		,					,	  18,	  2                                                                                  					'ǰ�� 
		ggoSpread.SSSetEdit 	C_ITEM_NM	,"ǰ���" 		,		25,		,					,	  40                                                                                                 					'ǰ��� 
		ggoSpread.SSSetEdit 	C_SPEC		,"ǰ��԰�"		,			20
		ggoSpread.SSSetEdit 	C_CUST_ITEM_CD	,"��ǰ��"		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_BP_ITEM_NM	,"��ǰ���"	,		25,		,					,	  40
		ggoSpread.SSSetEdit 	C_BP_ITEM_SPEC	,"��ǰ��԰�" 	,			20
		ggoSpread.SSSetFloat 	C_SO_QTY	,"����" 		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_SO_UNIT	,"����" 		,			8,		,					,	  3,	  2
		ggoSpread.SSSetFloat 	C_SO_PRICE	,"�ܰ�" 		,			15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_NET_AMT2	,"�ݾ�" 		,			15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetDate 	C_DLVY_DT	,"������" 		,		10,		2,					parent.gDateFormat
		ggoSpread.SSSetFloat 	C_VAT_AMT2	,"�ΰ����ݾ�" 	,		15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_VAT_RATE2	,"�ΰ�����"		,		10,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_VAT_TYPE2	,"�ΰ�������"	,		10,		,					,	  5,	  2
		ggoSpread.SSSetEdit 	C_VAT_TYPE_NM2	,"�ΰ���������" 	,	20
		ggoSpread.SSSetEdit 	C_VAT_INC_FLAG	,"�ΰ������Ա���" 	,20		'�ΰ������Ա��� 

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

		ggoSpread.SpreadLock 	C_CFM_FLAG		, -1, -1		'����Ȯ������ 
		ggoSpread.SpreadLock 	C_SOLD_TO_PARTY		, -1, -1		'���ֹ��� 
		ggoSpread.SpreadLock 	C_BP_FULL_NM		, -1, -1		'���ֹ��θ� 
		ggoSpread.SpreadLock 	C_CUST_PO_NO		, -1, -1		'�����ֹ�ȣ 
		ggoSpread.SpreadLock 	C_SO_NO			, -1, -1		'���ֹ�ȣ 
		ggoSpread.SpreadLock 	C_EXPORT_FLAG		, -1, -1		'�����ڱ��� 
		ggoSpread.SpreadLock 	C_SO_DT			, -1, -1		'������ 
		ggoSpread.SpreadLock 	C_SALES_GRP		, -1, -1		'�����׷� 
		ggoSpread.SpreadLock 	C_SALES_GRP_FULL_NM	, -1, -1		'�����׷�� 
		ggoSpread.SpreadLock 	C_CUR			, -1, -1		'ȭ�� 
		ggoSpread.SpreadLock 	C_NET_AMT		, -1, -1		'���ֱݾ� 
		ggoSpread.SpreadLock 	C_VAT_AMT		, -1, -1		'�ΰ����ݾ� 
		ggoSpread.SpreadLock 	C_NET_VAT_TOTAMT	, -1, -1		'�����ѱݾ� 
		ggoSpread.SpreadLock 	C_VAT_TYPE		, -1, -1		'�ΰ������� 
		ggoSpread.SpreadLock 	C_VAT_TYPE_NM		, -1, -1		'�ΰ��������� 
		ggoSpread.SpreadLock 	C_VAT_RATE		, -1, -1		'�ΰ����� 
		ggoSpread.SpreadLock 	C_PAY_METH		, -1, -1		'������� 
		ggoSpread.SpreadLock 	C_PAY_METH_NM		, -1, -1		'��������� 
		ggoSpread.SpreadLock 	C_INCOTERMS		, -1, -1		'�������� 
		ggoSpread.SpreadLock 	C_INCOTERMS_NM		, -1, -1		'�������Ǹ� 
		ggoSpread.SpreadLock 	C_HIDDEN_CFM_FLAG	, -1, -1	'����Ȯ������(HIDDEN)


		.ReDraw = True
	End With
End Sub

Sub SetSpreadLock2()
	With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2

		ggoSpread.SpreadLock	C_ITEM_CD	, -1, -1	'ǰ�� 
		ggoSpread.SpreadLock	C_ITEM_NM	, -1, -1	'ǰ��� 
		ggoSpread.SpreadLock	C_SPEC		, -1, -1	'ǰ��԰� 
		ggoSpread.SpreadLock	C_CUST_ITEM_CD	, -1, -1	'��ǰ�� 
		ggoSpread.SpreadLock	C_BP_ITEM_NM	, -1, -1	'��ǰ��� 
		ggoSpread.SpreadLock	C_BP_ITEM_SPEC	, -1, -1	'��ǰ��԰� 
		ggoSpread.SpreadLock	C_SO_QTY	, -1, -1	'���� 
		ggoSpread.SpreadLock	C_SO_UNIT	, -1, -1	'���� 
		ggoSpread.SpreadLock	C_SO_PRICE	, -1, -1	'�ܰ� 
		ggoSpread.SpreadLock	C_NET_AMT2	, -1, -1	'�ݾ� 
		ggoSpread.SpreadLock	C_DLVY_DT	, -1, -1	'������ 
		ggoSpread.SpreadLock	C_VAT_AMT2	, -1, -1	'�ΰ����ݾ� 
		ggoSpread.SpreadLock	C_VAT_RATE2	, -1, -1	'�ΰ����� 
		ggoSpread.SpreadLock	C_VAT_TYPE2	, -1, -1	'�ΰ������� 
		ggoSpread.SpreadLock	C_VAT_TYPE_NM2	, -1, -1	'�ΰ��������� 
		ggoSpread.SpreadLock	C_VAT_INC_FLAG	, -1, -1	'�ΰ������Ա��� 

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


		ggoSpread.SSSetProtected 	C_CFM_FLAG		, pvStartRow, pvEndRow		'����Ȯ������		ggoSpread.SSSetProtected 	C_SOLD_TO_PARTY		, pvStartRow, pvEndRow		'���ֹ��� 
		ggoSpread.SSSetProtected 	C_BP_FULL_NM		, pvStartRow, pvEndRow		'���ֹ��θ� 
		ggoSpread.SSSetProtected 	C_CUST_PO_NO		, pvStartRow, pvEndRow		'�����ֹ�ȣ 
		ggoSpread.SSSetProtected 	C_SO_NO			, pvStartRow, pvEndRow		'���ֹ�ȣ 
		ggoSpread.SSSetProtected 	C_EXPORT_FLAG		, pvStartRow, pvEndRow		'�����ڱ��� 
		ggoSpread.SSSetProtected 	C_SO_DT			, pvStartRow, pvEndRow		'������ 
		ggoSpread.SSSetProtected 	C_SALES_GRP		, pvStartRow, pvEndRow		'�����׷� 
		ggoSpread.SSSetProtected 	C_SALES_GRP_FULL_NM	, pvStartRow, pvEndRow		'�����׷�� 
		ggoSpread.SSSetProtected 	C_CUR			, pvStartRow, pvEndRow		'ȭ�� 
		ggoSpread.SSSetProtected 	C_NET_AMT		, pvStartRow, pvEndRow		'���ֱݾ� 
		ggoSpread.SSSetProtected 	C_VAT_AMT		, pvStartRow, pvEndRow		'�ΰ����ݾ� 
		ggoSpread.SSSetProtected 	C_NET_VAT_TOTAMT	, pvStartRow, pvEndRow		'�����ѱݾ� 
		ggoSpread.SSSetProtected 	C_VAT_TYPE		, pvStartRow, pvEndRow		'�ΰ������� 
		ggoSpread.SSSetProtected 	C_VAT_TYPE_NM		, pvStartRow, pvEndRow		'�ΰ��������� 
		ggoSpread.SSSetProtected 	C_VAT_RATE		, pvStartRow, pvEndRow		'�ΰ����� 
		ggoSpread.SSSetProtected 	C_PAY_METH		, pvStartRow, pvEndRow		'������� 
		ggoSpread.SSSetProtected 	C_PAY_METH_NM		, pvStartRow, pvEndRow		'��������� 
		ggoSpread.SSSetProtected 	C_INCOTERMS		, pvStartRow, pvEndRow		'�������� 
		ggoSpread.SSSetProtected 	C_INCOTERMS_NM		, pvStartRow, pvEndRow		'�������Ǹ� 
		ggoSpread.SSSetProtected 	C_HIDDEN_CFM_FLAG	, pvStartRow, pvEndRow		'����Ȯ������(HIDDEN)

		.ReDraw = True
	End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
	With frm1.vspdData2
		.ReDraw = False
		ggoSpread.Source = frm1.vspdData2



		ggoSpread.SSSetProtected	C_ITEM_CD	, pvStartRow, pvEndRow	'ǰ�� 
		ggoSpread.SSSetProtected	C_ITEM_NM	, pvStartRow, pvEndRow	'ǰ��� 
		ggoSpread.SSSetProtected	C_SPEC		, pvStartRow, pvEndRow	'ǰ��԰� 
		ggoSpread.SSSetProtected	C_CUST_ITEM_CD	, pvStartRow, pvEndRow	'��ǰ�� 
		ggoSpread.SSSetProtected	C_BP_ITEM_NM	, pvStartRow, pvEndRow	'��ǰ��� 
		ggoSpread.SSSetProtected	C_BP_ITEM_SPEC	, pvStartRow, pvEndRow	'��ǰ��԰� 
		ggoSpread.SSSetProtected	C_SO_QTY	, pvStartRow, pvEndRow	'���� 
		ggoSpread.SSSetProtected	C_SO_UNIT	, pvStartRow, pvEndRow	'���� 
		ggoSpread.SSSetProtected	C_SO_PRICE	, pvStartRow, pvEndRow	'�ܰ� 
		ggoSpread.SSSetProtected	C_NET_AMT2	, pvStartRow, pvEndRow	'�ݾ� 
		ggoSpread.SSSetProtected	C_DLVY_DT	, pvStartRow, pvEndRow	'������ 
		ggoSpread.SSSetProtected	C_VAT_AMT2	, pvStartRow, pvEndRow	'�ΰ����ݾ� 
		ggoSpread.SSSetProtected	C_VAT_RATE2	, pvStartRow, pvEndRow	'�ΰ����� 
		ggoSpread.SSSetProtected	C_VAT_TYPE2	, pvStartRow, pvEndRow	'�ΰ������� 
		ggoSpread.SSSetProtected	C_VAT_TYPE_NM2	, pvStartRow, pvEndRow	'�ΰ��������� 
		ggoSpread.SSSetProtected	C_VAT_INC_FLAG	, pvStartRow, pvEndRow	'�ΰ������Ա��� 

		.ReDraw = True
	End With
End Sub

'============================= 2.2.3 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_SEL_YN		= 1		'���� 
	C_CFM_FLAG		= 2		'����Ȯ������ 
	C_SOLD_TO_PARTY		= 3		'���ֹ��� 
	C_BP_FULL_NM		= 4		'���ֹ��θ� 
	C_CUST_PO_NO		= 5		'�����ֹ�ȣ 
	C_SO_NO			= 6		'���ֹ�ȣ 
	C_EXPORT_FLAG		= 7		'�����ڱ��� 
	C_SO_DT			= 8		'������ 
	C_SALES_GRP		= 9		'�����׷� 
	C_SALES_GRP_FULL_NM	= 10		'�����׷�� 
	C_CUR			= 11		'ȭ�� 
	C_NET_AMT		= 12		'���ֱݾ� 
	C_VAT_AMT		= 13		'�ΰ����ݾ� 
	C_NET_VAT_TOTAMT	= 14		'�����ѱݾ� 
	C_VAT_TYPE		= 15		'�ΰ������� 
	C_VAT_TYPE_NM		= 16		'�ΰ��������� 
	C_VAT_RATE		= 17		'�ΰ����� 
	C_PAY_METH		= 18		'������� 
	C_PAY_METH_NM		= 19		'��������� 
	C_INCOTERMS		= 20		'�������� 
	C_INCOTERMS_NM		= 21		'�������Ǹ� 
	C_HIDDEN_CFM_FLAG	= 22		'����Ȯ������(HIDDEN)
End Sub

Sub InitSpreadPosVariables2()
	C_ITEM_CD	= 1	'ǰ�� 
	C_ITEM_NM	= 2	'ǰ��� 
	C_SPEC		= 3	'ǰ��԰� 
	C_CUST_ITEM_CD	= 4	'��ǰ�� 
	C_BP_ITEM_NM	= 5	'��ǰ��� 
	C_BP_ITEM_SPEC	= 6	'��ǰ��԰� 
	C_SO_QTY	= 7	'���� 
	C_SO_UNIT	= 8	'���� 
	C_SO_PRICE	= 9	'�ܰ� 
	C_NET_AMT2	= 10	'�ݾ� 
	C_DLVY_DT	= 11	'������ 
	C_VAT_AMT2	= 12	'�ΰ����ݾ� 
	C_VAT_RATE2	= 13	'�ΰ����� 
	C_VAT_TYPE2	= 14	'�ΰ������� 
	C_VAT_TYPE_NM2	= 15	'�ΰ��������� 
	C_VAT_INC_FLAG	= 16	'�ΰ������Ա��� 
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
			C_SEL_YN		= iCurColumnPos(1)	'���� 
			C_CFM_FLAG		= iCurColumnPos(2)	'����Ȯ������ 
			C_SOLD_TO_PARTY		= iCurColumnPos(3)	'���ֹ��� 
			C_BP_FULL_NM		= iCurColumnPos(4)	'���ֹ��θ� 
			C_CUST_PO_NO		= iCurColumnPos(5)	'�����ֹ�ȣ 
			C_SO_NO			= iCurColumnPos(6)	'���ֹ�ȣ 
			C_EXPORT_FLAG		= iCurColumnPos(7)	'�����ڱ��� 
			C_SO_DT			= iCurColumnPos(8)	'������ 
			C_SALES_GRP		= iCurColumnPos(9)	'�����׷� 
			C_SALES_GRP_FULL_NM	= iCurColumnPos(10)	'�����׷�� 
			C_CUR			= iCurColumnPos(11)	'ȭ�� 
			C_NET_AMT		= iCurColumnPos(12)	'���ֱݾ� 
			C_VAT_AMT		= iCurColumnPos(13)	'�ΰ����ݾ� 
			C_NET_VAT_TOTAMT	= iCurColumnPos(14)	'�����ѱݾ� 
			C_VAT_TYPE		= iCurColumnPos(15)	'�ΰ������� 
			C_VAT_TYPE_NM		= iCurColumnPos(16)	'�ΰ��������� 
			C_VAT_RATE		= iCurColumnPos(17)	'�ΰ����� 
			C_PAY_METH		= iCurColumnPos(18)	'������� 
			C_PAY_METH_NM		= iCurColumnPos(19)	'��������� 
			C_INCOTERMS		= iCurColumnPos(20)	'�������� 
			C_INCOTERMS_NM		= iCurColumnPos(21)	'�������Ǹ� 
			C_HIDDEN_CFM_FLAG	= iCurColumnPos(22)	'����Ȯ������(HIDDEN)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
            		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)


			C_ITEM_CD	= iCurColumnPos(1)	'ǰ�� 
			C_ITEM_NM	= iCurColumnPos(2)	'ǰ��� 
			C_SPEC		= iCurColumnPos(3)	'ǰ��԰� 
			C_CUST_ITEM_CD	= iCurColumnPos(4)	'��ǰ�� 
			C_BP_ITEM_NM	= iCurColumnPos(5)	'��ǰ��� 
			C_BP_ITEM_SPEC	= iCurColumnPos(6)	'��ǰ��԰� 
			C_SO_QTY	= iCurColumnPos(7)	'���� 
			C_SO_UNIT	= iCurColumnPos(8)	'���� 
			C_SO_PRICE	= iCurColumnPos(9)	'�ܰ� 
			C_NET_AMT2	= iCurColumnPos(10)	'�ݾ� 
			C_DLVY_DT	= iCurColumnPos(11)	'������ 
			C_VAT_AMT2	= iCurColumnPos(12)	'�ΰ����ݾ� 
			C_VAT_RATE2	= iCurColumnPos(13)	'�ΰ����� 
			C_VAT_TYPE2	= iCurColumnPos(14)	'�ΰ������� 
			C_VAT_TYPE_NM2	= iCurColumnPos(15)	'�ΰ��������� 
			C_VAT_INC_FLAG	= iCurColumnPos(16)	'�ΰ������Ա��� 

	End Select
End Sub

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : ���Ÿ� ���� �׸����� ���� �κ��� ����ȸ� �� �Լ��� ���� �ؾ���.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
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
	If y<20 Then			'2003-03-01 Release �߰� 
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
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###�׸��� ������ ���Ǻκ�###

 	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("0101111111")         'ȭ�麰 ���� 
	End If

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		lgSpdHdrClicked = 0		'2003-03-01 Release �߰� 
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
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
 	Dim strShowDataFirstRow
 	Dim strShowDataLastRow
 	Dim i,k
 	Dim strFlag,strFlag1

 	gMouseClickStatus = "SP2C"

 	Set gActiveSpdSheet = frm1.vspdData2

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
	End If

 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If


End Sub

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
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
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
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
'		Cookie ����Լ� 
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
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()	'###�׸��� ������ ���Ǻκ�###

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
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)

End Sub

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : vspdData2_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : �׸��� �������� �Ѵ�.
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
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
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

		lRow = frm1.vspdData.ActiveRow	'###�׸��� ������ ���Ǻκ�###
		frm1.vspdData2.Redraw = True
    End If

 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	'###�׸��� ������ ���Ǻκ�###
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release �߰� 
		Exit Sub
	End If

	'/* 9�� ������ġ : ������ Ű���� �Է��� ä �ٸ� ��������� �ű��� ���ϵ��� �������� ���� �߰� - START */
	Dim lRow
	'/* 9�� ������ġ : ������ Ű���� �Է��� ä �ٸ� ��������� �ű��� ���ϵ��� �������� ���� �߰� - END */

	Set gActiveSpdSheet = frm1.vspdData

	frm1.vspdData.redraw = false
	If Row <> NewRow And NewRow > 0 Then
		With frm1
			.vspdData.redraw = false
			'/* 8�� ������ġ : ���� �������忡 �ʼ��Է� �ʵ� üũ - START */
		'	If DefaultCheck = False Then
		'		.vspdData.Row = Row
		'		.vspdData.Col = 1
		'		.vspdData2.focus
    	'		Exit Sub
		'	End If
			'/* 8�� ������ġ : ���� �������忡 �ʼ��Է� �ʵ� üũ - END */

			'/* 9�� ������ġ: '�ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - START */
			If CheckRunningBizProcess = True Then
				.vspdData.Row = Row
				.vspdData.Col = 1
				Exit Sub
			End If
			'/* 9�� ������ġ: '�ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - END */
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
	'/��û���� ����Ǹ� ��η��� �����Ѵ�.(��û�� * ��κ���)
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

    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
    '/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - END */
		If lgPageNo <> "" Then			'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
    	'/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - START */
    	If ShowDataLastRow < NewTop + VisibleRowCnt(.vspdData2, NewTop) Then	        '��: ������ üũ 
		'/* 9�� ������ġ: �ػ󵵿� ������� �������ǵ��� ���� - END */
'    		If lgStrPrevKeyM(lRow - 1) <> "" Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
    		If lgPageNo1 <> "" Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================
'==========================================================================================
'   Event Name : txtFrDt
'   Event Desc : �������� 
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
'   Event Desc : �������� 
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
'   Event Desc : ������ 
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
'   Event Desc : ������ 
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
Function FncQuery() '###�׸��� ������ ���Ǻκ�###
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

	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
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
			Call DisplayMsgBox("17a003", "X","��������", "X")
			Exit Function
		End If

	End with

	'-----------------------
   	'Query function call area
    	'-----------------------
	If DbQuery = False then
		Exit Function
	End If																		'��: Query db data

	Set gActiveElement = document.activeElement

    	FncQuery = True
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave()
    FncSave = False

    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>

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
    If Not ggoSpread.SSDefaultCheck Then    'Not chkField(Document, "2") OR      '��: Check contents area
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data

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
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			'��: "Will you destory previous data"
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

			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '���ֹ����ڵ� 
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)			'�������� From
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)			'�������� To
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'������ From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'������ To

			if .rdoCfmFlag(0).checked = true Then					'����Ȯ������ 
				strVal = strVal & "&rdoCfmFlag=" & "Y"	'Ȯ�� 
			else
				strVal = strVal & "&rdoCfmFlag=" & "N"	'��Ȯ�� 
			End if

			strVal = strVal & "&txtPO_NO=" & Trim(.txtPO_NO.value)			'�����ֹ�ȣ 
			strVal = strVal & "&txtSO_NO=" & Trim(.txtSO_NO.value)			'���ֹ�ȣ 

			strVal = strVal & "&lgPageNo=" & lgPageNo                  		'��: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	Else
	        	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '���ֹ����ڵ� 
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)			'�������� From
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)			'�������� To
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'������ From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'������ To

			if .rdoCfmFlag(0).checked = true Then					'����Ȯ������ 
				strVal = strVal & "&rdoCfmFlag=" & "Y"	'Ȯ�� 
			else
				strVal = strVal & "&rdoCfmFlag=" & "N"	'��Ȯ�� 
			End if

			strVal = strVal & "&txtPO_NO=" & Trim(.txtPO_NO.value)			'�����ֹ�ȣ 
			strVal = strVal & "&txtSO_NO=" & Trim(.txtSO_NO.value)			'���ֹ�ȣ 

			strVal = strVal & "&lgPageNo=" & lgPageNo                  		'��: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	End If
	End with



	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 

	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)
	DbQueryOk = False

	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows

	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field

	if frm1.rdoCfmFlag(0).checked = true Then	'Ȯ�����¶�� 
		Call SetToolBar("11001000000011")				'��ư ���� ���� 
	Else						'��Ȯ�����¶�� 
		Call SetToolBar("11001011000011")				'��ư ���� ���� 
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
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
			Else
				TmpArrPrevKey=lgStrPrevKeyM
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lgStrPrevKeyM(intTRow+intARow - 1)
				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
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

		.vspdData.Col = C_SO_NO		'���ֹ�ȣ 
		strSO_NO  = .vspdData.Text

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows

		strVal = strVal & "&strSO_NO=" & trim(strSO_NO)

		strVal = strVal & "&lgStrPrevKeyM="  & lgStrPrevKeyM(Row - 1)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'��: Next key tag
		strVal = strVal & "&lglngHiddenRows=" & .vspdData.MaxRows

		.vspdData.redraw = True

	End With


	Call RunMyBizASP(MyBizASP, strVal)
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
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
	'On Error Resume Next                                                   <%'��: Protect system from crashing%>

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
	' Data ���� ��Ģ 
	' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ 

	For lRow = 1 To .vspdData.MaxRows

		.vspdData.Row = lRow
		.vspdData.Col = 0

		Select Case .vspdData.Text
		    Case ggoSpread.DeleteFlag													'��: ���� 
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep
					.vspdData.Row		= lRow
					.vspdData.Col		= C_SO_NO
					istrSO_NO	= Trim(.vspdData.text)
					strDel = strDel & Trim(istrSO_NO)		& parent.gColSep	'���ֹ�ȣ 

					.vspdData.Col		= C_HIDDEN_CFM_FLAG
					strHIDDEN_CFM_FLAG	= Trim(.vspdData.text)
					strDel = strDel & Trim(strHIDDEN_CFM_FLAG)		& parent.gColSep	'���ֹ�ȣ 
										'��: U=Update
		                        strDel = strDel & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1

		    Case ggoSpread.UpdateFlag													'��: ���� 
					strDel = strDel & "U" & parent.gColSep & lRow & parent.gColSep
					.vspdData.Row		= lRow
					.vspdData.Col		= C_SO_NO
					istrSO_NO	= Trim(.vspdData.text)
					strDel = strDel & Trim(istrSO_NO)		& parent.gColSep	'���ֹ�ȣ 

					.vspdData.Col		= C_HIDDEN_CFM_FLAG
					strHIDDEN_CFM_FLAG	= Trim(.vspdData.text)
					strDel = strDel & Trim(strHIDDEN_CFM_FLAG)		& parent.gColSep	'���ֹ�ȣ 
										'��: U=Update
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


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'��: �����Ͻ� ASP �� ���� %>

	End With




    DbSave = True
End Function

'========================================================================================================
Function FncDeleteRow()
	Dim lDelRows, lDelRow

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncDeleteRow = False                                                          '��: Processing is NG

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
       FncDeleteRow = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'========================================================================================================
Function FncDelete()
    Dim intRetCD

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncDelete = False                                                             '��: Processing is NG

    '------ Developer Coding part (Start ) --------------------------------------------------------------
    '�ʿ������???
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If

    Call ggoOper.ClearField(Document, "A")

    '------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncDelete = True                                                           '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function



'========================================================================================================
Function FncCancel()
	Dim iDx

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status

    FncCancel = False                                                             '��: Processing is NG

	If frm1.vspdData.MaxRows < 1 Then Exit Function
    ggoSpread.Source = Frm1.vspdData

    Call CancelSum()

    ggoSpread.EditUndo
	'------ Developer Coding part (Start ) --------------------------------------------------------------

    '------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncCancel = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

End Function


'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
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

	arrParam(0) = "���ֹ���"
	arrParam(1) = "B_Biz_Partner"

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
'	arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE In ('C','CS') And IN_OUT_FLAG = 'O'"
	arrParam(5) = "���ֹ���"

    	arrField(0) = "BP_Cd"
    	arrField(1) = "BP_NM"

    	arrHeader(0) = "���ֹ���"
    	arrHeader(1) = "���ֹ��θ�"

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
'   Event Desc : �ϰ����� ��ư�� Ŭ���� ��� �߻� 
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
'   Event Desc : �ϰ�������� ��ư�� Ŭ���� ��� �߻� 
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
'   Event Desc : Ȯ��ó�� �Ǵ� Ȯ����� ��ư�� Ŭ���� ��� �߻� 
'==========================================================================================
Sub btnProcessCfm_OnClick()
	Dim lRow
	Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel, strTxt, k
	Dim istrSO_NO
	Dim strHIDDEN_CFM_FLAG


	Call LayerShowHide(1)
	'On Error Resume Next                                                   <%'��: Protect system from crashing%>

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
	' Data ���� ��Ģ 
	' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ 

	For lRow = 1 To .vspdData.MaxRows

		.vspdData.Row = lRow
		.vspdData.Col = 0

		Select Case .vspdData.Text
		    Case ggoSpread.UpdateFlag													'��: ���� 
					strDel = strDel & "U" & parent.gColSep & lRow & parent.gColSep
					.vspdData.Row		= lRow
					.vspdData.Col		= C_SO_NO
					istrSO_NO	= Trim(.vspdData.text)
					strDel = strDel & Trim(istrSO_NO)		& parent.gColSep	'���ֹ�ȣ 

					.vspdData.Col		= C_HIDDEN_CFM_FLAG
					strHIDDEN_CFM_FLAG	= Trim(.vspdData.text)
					strDel = strDel & Trim(strHIDDEN_CFM_FLAG)		& parent.gColSep	'���ֹ�ȣ 
										'��: U=Update
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


	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'��: �����Ͻ� ASP �� ���� %>

	End With

End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : �ϰ�Ȯ�� ��ư�� Ŭ���� ��� �߻� 
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
'   Event Desc : �ϰ�Ȯ����� ��ư�� Ŭ���� ��� �߻� 
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
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��Ƽ���۴� ����Ȯ��/����</font></td>
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
								<TD CLASS="TD5" NOWRAP>���ֹ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtSupplierCd"  SIZE=10 MAXLENGTH=10 ALT="���ֹ���"  tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
										       <INPUT TYPE=TEXT Name="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="���ֹ���"  tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>��������</TD>
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
								<TD CLASS="TD5" NOWRAP>������</TD>
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
								<TD CLASS="TD5" NOWRAP>����Ȯ��ó������</TD>
								<TD CLASS="TD6" NOWRAP>
									<input type=radio CLASS = "RADIO" name="rdoCfmFlag" id="rdoCfmFlagN" value="Y" tag = "11" checked>
										<label for="rdoCfmFlagN">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS="RADIO" name="rdoCfmFlag" id="rdoCfmFlagY" value="N" tag = "11" >
										<label for="rdoCfmFlagY">��Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ֹ�ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtPO_NO" SIZE=29 MAXLENGTH=18  tag="11" ALT="�����ֹ�ȣ" STYLE="text-transform:uppercase"></TD>
								<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtSO_NO" SIZE=29 MAXLENGTH=18 tag="11" ALT="���ֹ�ȣ" STYLE="text-transform:uppercase"></TD>

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
				<BUTTON name="btnSelect" class="clsmbtn" >�ϰ�����</button>&nbsp;
				<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">�ϰ��������</BUTTON>&nbsp;&nbsp;
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
