<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm112qa1
'*  4. Program Name         : ��Ƽ���۴ϼ�����������ȸ(���ֺ�)
'*  5. Program Desc         : ��Ƽ���۴ϼ�����������ȸ(���ֺ�)-��Ƽ 
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "ksm112qb1.asp"
											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_PO_COMPANY	'���ֹ��� 
Dim C_PO_COMPANY_NM	'���ֹ��θ� 
Dim C_SO_NO		'���ֹ�ȣ 
Dim C_SO_SEQ_NO		'���ּ��� 
Dim C_ITEM_CD		'ǰ�� 
Dim C_ITEM_NM		'ǰ��� 
Dim C_SPEC		'ǰ��԰� 
Dim C_PO_STS		'���ֹ��λ��� 
Dim C_SO_STS		'���ֹ��λ��� 
Dim C_UNIT		'���� 
Dim C_PO_QTY		'���ּ��� 
Dim C_SO_QTY		'���ּ��� 
Dim C_PO_LC_QTY		'����L/C���� 
Dim C_SO_LC_QTY		'����L/C���� 
Dim C_SO_REQ_QTY	'���Ͽ�û���� 
Dim C_SO_ISSUE_QTY	'������ 
Dim C_SO_CC_QTY		'����������� 
Dim C_PO_CC_QTY		'����������� 
Dim C_PO_RCPT_QTY	'�԰���� 
Dim C_SO_BILL_QTY	'������� 
Dim C_PO_IV_QTY		'���Լ��� 
Dim C_PO_NO		'���ֹ���ȣ 
Dim C_PO_SEQ_NO		'���� 
Dim C_BP_ITEM_CD	'��ǰ�� 
Dim C_BP_ITEM_NM	'��ǰ��� 


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
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
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

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
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

		ggoSpread.SSSetEdit 	C_PO_COMPANY	        	,"���ֹ���"		,15	'���ֹ��� 
		ggoSpread.SSSetEdit 	C_PO_COMPANY_NM	        	,"���ֹ��θ�"		,25	'���ֹ��� 
		ggoSpread.SSSetEdit 	C_SO_NO		        	,"���ֹ�ȣ"		,15	'���ֹ�ȣ 
		ggoSpread.SSSetEdit 	C_SO_SEQ_NO			,"���ּ���"		,15	'���ּ��� 
		ggoSpread.SSSetEdit 	C_ITEM_CD			,"ǰ��"		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_ITEM_NM			,"ǰ���"		,		25,		,					,	  40
		ggoSpread.SSSetEdit 	C_SPEC		        	,"ǰ��԰�"		,			20
		ggoSpread.SSSetEdit 	C_PO_STS			,"���ֹ��λ���"	,20	'���ֹ��λ��� 
		ggoSpread.SSSetEdit 	C_SO_STS			,"���ֹ��λ���"	,20	'���ֹ��λ��� 
		ggoSpread.SSSetEdit 	C_UNIT		        	,"����"		,			8,		,					,	  3,	  2
		ggoSpread.SSSetFloat 	C_PO_QTY			,"���ּ���"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_QTY			,"���ּ���"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_LC_QTY			,"����L/C����"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_LC_QTY			,"����L/C����"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_REQ_QTY	        	,"���Ͽ�û����"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_ISSUE_QTY	        	,"������"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_CC_QTY			,"�����������"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_CC_QTY			,"�����������"	,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_RCPT_QTY	        	,"�԰����"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_SO_BILL_QTY	        	,"�������"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat 	C_PO_IV_QTY			,"���Լ���"		,			15,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit 	C_PO_NO		        	,"���ֹ���ȣ"		,15	'���ֹ���ȣ 
		ggoSpread.SSSetEdit 	C_PO_SEQ_NO			,"����"		,15	'���� 
		ggoSpread.SSSetEdit 	C_BP_ITEM_CD	        	,"��ǰ��"		,			18,		,					,	  18,	  2
		ggoSpread.SSSetEdit 	C_BP_ITEM_NM	        	,"��ǰ���"	,		25,		,					,	  40


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


		ggoSpread.SpreadLock 	C_PO_COMPANY	        , -1, -1	'���ֹ��� 
		ggoSpread.SpreadLock 	C_PO_COMPANY_NM	        , -1, -1	'���ֹ��θ� 
		ggoSpread.SpreadLock 	C_SO_NO		        , -1, -1	'���ֹ�ȣ 
		ggoSpread.SpreadLock 	C_SO_SEQ_NO		, -1, -1	'���ּ��� 
		ggoSpread.SpreadLock 	C_ITEM_CD		, -1, -1	'ǰ�� 
		ggoSpread.SpreadLock 	C_ITEM_NM		, -1, -1	'ǰ��� 
		ggoSpread.SpreadLock 	C_SPEC		        , -1, -1	'ǰ��԰� 
		ggoSpread.SpreadLock 	C_PO_STS		, -1, -1	'���ֹ��λ��� 
		ggoSpread.SpreadLock 	C_SO_STS		, -1, -1	'���ֹ��λ��� 
		ggoSpread.SpreadLock 	C_UNIT		        , -1, -1	'���� 
		ggoSpread.SpreadLock 	C_PO_QTY		, -1, -1	'���ּ��� 
		ggoSpread.SpreadLock 	C_SO_QTY		, -1, -1	'���ּ��� 
		ggoSpread.SpreadLock 	C_PO_LC_QTY		, -1, -1	'����L/C���� 
		ggoSpread.SpreadLock 	C_SO_LC_QTY		, -1, -1	'����L/C���� 
		ggoSpread.SpreadLock 	C_SO_REQ_QTY	        , -1, -1	'���Ͽ�û���� 
		ggoSpread.SpreadLock 	C_SO_ISSUE_QTY	        , -1, -1	'������ 
		ggoSpread.SpreadLock 	C_SO_CC_QTY		, -1, -1	'����������� 
		ggoSpread.SpreadLock 	C_PO_CC_QTY		, -1, -1	'����������� 
		ggoSpread.SpreadLock 	C_PO_RCPT_QTY	        , -1, -1	'�԰���� 
		ggoSpread.SpreadLock 	C_SO_BILL_QTY	        , -1, -1	'������� 
		ggoSpread.SpreadLock 	C_PO_IV_QTY		, -1, -1	'���Լ��� 
		ggoSpread.SpreadLock 	C_PO_NO		        , -1, -1	'���ֹ���ȣ 
		ggoSpread.SpreadLock 	C_PO_SEQ_NO		, -1, -1	'���� 
		ggoSpread.SpreadLock 	C_BP_ITEM_CD	        , -1, -1	'��ǰ�� 
		ggoSpread.SpreadLock 	C_BP_ITEM_NM	        , -1, -1	'��ǰ��� 


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

		ggoSpread.SSSetProtected 	C_PO_COMPANY	        , pvStartRow, pvEndRow		'���ֹ��� 
		ggoSpread.SSSetProtected 	C_PO_COMPANY_NM	        , pvStartRow, pvEndRow		'���ֹ��θ� 
		ggoSpread.SSSetProtected 	C_SO_NO		        , pvStartRow, pvEndRow		'���ֹ�ȣ 
		ggoSpread.SSSetProtected 	C_SO_SEQ_NO		, pvStartRow, pvEndRow		'���ּ��� 
		ggoSpread.SSSetProtected 	C_ITEM_CD		, pvStartRow, pvEndRow		'ǰ�� 
		ggoSpread.SSSetProtected 	C_ITEM_NM		, pvStartRow, pvEndRow		'ǰ��� 
		ggoSpread.SSSetProtected 	C_SPEC		        , pvStartRow, pvEndRow		'ǰ��԰� 
		ggoSpread.SSSetProtected 	C_PO_STS		, pvStartRow, pvEndRow		'���ֹ��λ��� 
		ggoSpread.SSSetProtected 	C_SO_STS		, pvStartRow, pvEndRow		'���ֹ��λ��� 
		ggoSpread.SSSetProtected 	C_UNIT		        , pvStartRow, pvEndRow		'���� 
		ggoSpread.SSSetProtected 	C_PO_QTY		, pvStartRow, pvEndRow		'���ּ��� 
		ggoSpread.SSSetProtected 	C_SO_QTY		, pvStartRow, pvEndRow		'���ּ��� 
		ggoSpread.SSSetProtected 	C_PO_LC_QTY		, pvStartRow, pvEndRow		'����L/C���� 
		ggoSpread.SSSetProtected 	C_SO_LC_QTY		, pvStartRow, pvEndRow		'����L/C���� 
		ggoSpread.SSSetProtected 	C_SO_REQ_QTY	        , pvStartRow, pvEndRow		'���Ͽ�û���� 
		ggoSpread.SSSetProtected 	C_SO_ISSUE_QTY	        , pvStartRow, pvEndRow		'������ 
		ggoSpread.SSSetProtected 	C_SO_CC_QTY		, pvStartRow, pvEndRow		'����������� 
		ggoSpread.SSSetProtected 	C_PO_CC_QTY		, pvStartRow, pvEndRow		'����������� 
		ggoSpread.SSSetProtected 	C_PO_RCPT_QTY	        , pvStartRow, pvEndRow		'�԰���� 
		ggoSpread.SSSetProtected 	C_SO_BILL_QTY	        , pvStartRow, pvEndRow		'������� 
		ggoSpread.SSSetProtected 	C_PO_IV_QTY		, pvStartRow, pvEndRow		'���Լ��� 
		ggoSpread.SSSetProtected 	C_PO_NO		        , pvStartRow, pvEndRow		'���ֹ���ȣ 
		ggoSpread.SSSetProtected 	C_PO_SEQ_NO		, pvStartRow, pvEndRow		'���� 
		ggoSpread.SSSetProtected 	C_BP_ITEM_CD	        , pvStartRow, pvEndRow		'��ǰ�� 
		ggoSpread.SSSetProtected 	C_BP_ITEM_NM	        , pvStartRow, pvEndRow		'��ǰ��� 

		.ReDraw = True
	End With
End Sub



'============================= 2.2.3 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_PO_COMPANY	        = 1	'���ֹ��� 
	C_PO_COMPANY_NM	        = 2	'���ֹ��θ� 
	C_SO_NO		        = 3	'���ֹ�ȣ 
	C_SO_SEQ_NO		= 4	'���ּ��� 
	C_ITEM_CD		= 5	'ǰ�� 
	C_ITEM_NM		= 6	'ǰ��� 
	C_SPEC		        = 7	'ǰ��԰� 
	C_PO_STS		= 8	'���ֹ��λ��� 
	C_SO_STS		= 9	'���ֹ��λ��� 
	C_UNIT		        = 10	'���� 
	C_PO_QTY		= 11	'���ּ��� 
	C_SO_QTY		= 12	'���ּ��� 
	C_PO_LC_QTY		= 13	'����L/C���� 
	C_SO_LC_QTY		= 14	'����L/C���� 
	C_SO_REQ_QTY	        = 15	'���Ͽ�û���� 
	C_SO_ISSUE_QTY	        = 16	'������ 
	C_SO_CC_QTY		= 17	'����������� 
	C_PO_CC_QTY		= 18	'����������� 
	C_PO_RCPT_QTY	        = 19	'�԰���� 
	C_SO_BILL_QTY	        = 20	'������� 
	C_PO_IV_QTY		= 21	'���Լ��� 
	C_PO_NO		        = 22	'���ֹ���ȣ 
	C_PO_SEQ_NO		= 23	'���� 
	C_BP_ITEM_CD	        = 24	'��ǰ�� 
	C_BP_ITEM_NM	        = 25	'��ǰ��� 
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

			C_PO_COMPANY	        = iCurColumnPos(1)	'���ֹ��� 
			C_PO_COMPANY_NM	        = iCurColumnPos(2)	'���ֹ��θ� 
			C_SO_NO		        = iCurColumnPos(3)	'���ֹ�ȣ 
			C_SO_SEQ_NO		= iCurColumnPos(4)	'���ּ��� 
			C_ITEM_CD		= iCurColumnPos(5)	'ǰ�� 
			C_ITEM_NM		= iCurColumnPos(6)	'ǰ��� 
			C_SPEC		        = iCurColumnPos(7)	'ǰ��԰� 
			C_PO_STS		= iCurColumnPos(8)	'���ֹ��λ��� 
			C_SO_STS		= iCurColumnPos(9)	'���ֹ��λ��� 
			C_UNIT		        = iCurColumnPos(10)	'���� 
			C_PO_QTY		= iCurColumnPos(11)	'���ּ��� 
			C_SO_QTY		= iCurColumnPos(12)	'���ּ��� 
			C_PO_LC_QTY		= iCurColumnPos(13)	'����L/C���� 
			C_SO_LC_QTY		= iCurColumnPos(14)	'����L/C���� 
			C_SO_REQ_QTY	        = iCurColumnPos(15)	'���Ͽ�û���� 
			C_SO_ISSUE_QTY	        = iCurColumnPos(16)	'������ 
			C_SO_CC_QTY		= iCurColumnPos(17)	'����������� 
			C_PO_CC_QTY		= iCurColumnPos(18)	'����������� 
			C_PO_RCPT_QTY	        = iCurColumnPos(19)	'�԰���� 
			C_SO_BILL_QTY	        = iCurColumnPos(20)	'������� 
			C_PO_IV_QTY		= iCurColumnPos(21)	'���Լ��� 
			C_PO_NO		        = iCurColumnPos(22)	'���ֹ���ȣ 
			C_PO_SEQ_NO		= iCurColumnPos(23)	'���� 
			C_BP_ITEM_CD	        = iCurColumnPos(24)	'��ǰ�� 
			C_BP_ITEM_NM	        = iCurColumnPos(25)	'��ǰ��� 

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
	If y<20 Then			'2003-03-01 Release �߰� 
	    lgSpdHdrClicked = 1
	End If

    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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

	Call InitVariables

	Call SetDefaultVal

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
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
 Sub txtSo_Frdt_DblClick(Button)
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
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
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

			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.value)        '���ֹ����ڵ� 
			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '���ֹ����ڵ� 
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'������ From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'������ To

			if .rdoPostFlag2(0).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & ""
			elseif .rdoPostFlag2(1).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "N"
			else
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			End if

			strVal = strVal & "&lgPageNo=" & lgPageNo                  		'��: Next key tag
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    	Else
	        	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSpplCd=" & Trim(.txtSpplCd.value)        '���ֹ����ڵ� 
			strVal = strVal & "&txtSupplierCd=" & Trim(.txtSupplierCd.value)        '���ֹ����ڵ� 
			strVal = strVal & "&txtSo_Frdt=" & Trim(.txtSo_Frdt.text)		'������ From
			strVal = strVal & "&txtSo_Todt=" & Trim(.txtSo_Todt.text)		'������ To

			if .rdoPostFlag2(0).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & ""
			elseif .rdoPostFlag2(1).checked = true then
				strVal = strVal & "&rdoPostFlag2=" & "N"
			else
				strVal = strVal & "&rdoPostFlag2=" & "Y"
			End if

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
	Call SetToolBar("11000000000011")				'��ư ���� ���� 

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
			Case ggoSpread.UpdateFlag			'��: ����, �ű� 
				strVal = strVal & lRow & parent.gColSep	'��: U=Update

				'��� �������� 
				.vspdData.Col =C_CFM_YN		 	'Ȯ������ 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col =C_PO_COMPANY    		'���ֹ��� 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col =C_SO_COMPANY       	'���ֹ��� 
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
				.vspdData.Col =C_PO_NO              	'���ֹ�ȣ(���ֹ���ȣ)
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
'   Event Desc : �ϰ����� ��ư�� Ŭ���� ��� �߻� 
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
'   Event Desc : �ϰ�������� ��ư�� Ŭ���� ��� �߻� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��Ƽ���۴� ������������ȸ</font></td>
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
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtSpplCd"  SIZE=10 MAXLENGTH=10 ALT="���ֹ���"  tag="14X">
										       <INPUT TYPE=TEXT Name="txtSpplNm" SIZE=20 MAXLENGTH=18 ALT="���ֹ���"  tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>���ֹ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Name="txtSupplierCd"  SIZE=10 MAXLENGTH=10 ALT="���ֹ���"  tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
										       <INPUT TYPE=TEXT Name="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="���ֹ���"  tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������</TD>
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
								<TD CLASS="TD5" NOWRAP>������ó������</TD>
								<TD CLASS="TD6" NOWRAP>
										<input type=radio CLASS = "RADIO" name="rdoPostFlag2" id="rdoPostFlag" value="" tag = "11" checked>
											<label for="rdoPostFlag">��ü</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoPostFlag2" id="rdoPostFlagN" value="N" tag = "11" >
											<label for="rdoPostFlagN">����</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostFlag2" id="rdoPostFlagY" value="Y" tag = "11" >
											<label for="rdoPostFlagY">����</label>
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
