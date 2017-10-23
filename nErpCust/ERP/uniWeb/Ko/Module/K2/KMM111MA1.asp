<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : MM111MA1
'*  4. Program Name         : ��Ƽ���۴ϸ��Ե�� 
'*  5. Program Desc         : ��Ƽ���۴ϸ��Ե�� 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2005/03/08
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : MJG
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

Const BIZ_PGM_ID = "KMM111MB1.asp"
Const BIZ_PGM_ID2 = "KMM111MB101.asp"
'Const BIZ_PGM_SAVE_ID = "m2111mb5.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL = "KMM111QA1"
											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��� �������� 
Dim	C_Check								'���� 
Dim	C_CfmFlg								'Ȯ������ 
Dim	C_BlNo									'����ó���ݰ�꼭��ȣ 
Dim	C_BlCur									'ȭ�� 
Dim	C_BlDocAmt								'���ް��� 
Dim	C_BlVatDocAmt							'�ΰ����ݾ� 
Dim	C_BlTotDocAmt							'�հ�ݾ� 
Dim	C_BlVatType								'�ΰ������� 
Dim	C_BlVatTypeNm							'�ΰ��������� 
Dim	C_BlVatRt								'�ΰ����� 
Dim	C_BlPayMeth								'������� 
Dim C_BlPayMethNm							'�������� 
Dim	C_SoNo									'���ֹ�ȣ 
Dim	C_CustPoNo								'���ֹ�ȣ 
Dim C_PoCompanyCd
Dim C_SoCompanyCd
Dim C_TaxQty

'�ϴ� �������� 
Dim C_BillNo
Dim C_BillSeqNo

Dim C_CustItemCd							'ǰ�� 
Dim C_CustItemNm							'ǰ��� 
Dim C_CustItemSpec							'ǰ��԰� 
Dim C_BillQty								'���� 
Dim C_BillUnit								'���� 
Dim C_BillPrc								'�ܰ� 
Dim C_BillDocAmt							'�ݾ� 
Dim C_VatDocAmt								'�ΰ����ݾ� 
Dim C_VatType								'�ΰ������� 
Dim C_VatTypeNm								'�ΰ��������� 
Dim C_VatRate								'�ΰ����� 
Dim C_VatInc								'�ΰ������Ա��� 

Dim C_ParentRowNo							'���� row ��ȣ 
Dim C_Flag									'�ڱ� ��ȣ 



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
Dim lgCurrRow
Dim strInspClass

Dim lgPageNo1

Dim EndDate, StartDate,CurrDate, iDBSYSDate

' === 2005.07.22 ���� ===========================================================
StartDate   = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
StartDate   = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate     = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
' === 2005.07.22 ���� ===========================================================

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

    '###�˻�з��� ����κ� Start###
    strInspClass = "R"
	'###�˻�з��� ����κ� End###
    'ggoSpread.ClearSpreadData

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
	frm1.txtPurGrp.value = Parent.gPurGrp
	frm1.hdnUsrId.value = parent.gUsrID

	frm1.txtFrBillDt.Text=StartDate
	frm1.txtToBillDt.Text=EndDate
	frm1.txtBillDt.Text=EndDate
	frm1.txtPayDt.Text=EndDate

	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
	frm1.btnConfirm.disabled = True
	frm1.btnCancel.disabled = True


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

    .MaxCols = C_TaxQty + 1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")
		ggoSpread.SSSetCheck    C_Check					, "����"					, 8,,,true
		ggoSpread.SSSetCheck    C_CfmFlg					, "Ȯ������"				, 10,,,true
		ggoSpread.SSSetEdit		C_BlNo						, "����ó���ݰ�꼭��ȣ"	, 20
		ggoSpread.SSSetEdit		C_BlCur						, "ȭ��"					, 10
		ggoSpread.SSSetFloat	C_BlDocAmt					, "���ް���"				, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat	C_BlVatDocAmt				, "�ΰ����ݾ�"				, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat	C_BlTotDocAmt				, "�հ�ݾ�"				, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_BlVatType					, "�ΰ�������"				, 18
		ggoSpread.SSSetEdit		C_BlVatTypeNm				, "�ΰ���������"			, 18
		ggoSpread.SSSetFloat	C_BlVatRt					, "�ΰ�����"				, 12,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_BlPayMeth					, "�������"				, 18
		ggoSpread.SSSetEdit		C_BlPayMethNm				, "��������"				, 18
		ggoSpread.SSSetEdit		C_SoNo						, "���ֹ�ȣ"				, 18
		ggoSpread.SSSetEdit		C_CustPoNo					, "���ֹ�ȣ"				, 18
		ggoSpread.SSSetEdit		C_PoCompanyCd				, "���ֹ���"				, 18
		ggoSpread.SSSetEdit		C_SoCompanyCd				, "���ֹ���"				, 18
		ggoSpread.SSSetFloat	C_TaxQty					, "����"				, 12,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"

'	Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)
	Call ggoSpread.SSSetColHidden(C_PoCompanyCd,	C_PoCompanyCd,	True)
	Call ggoSpread.SSSetColHidden(C_SoCompanyCd,	C_SoCompanyCd,	True)
	Call ggoSpread.SSSetColHidden(C_TaxQty,	C_TaxQty,	True)

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


		ggoSpread.SSSetEdit 	C_BillNo					, "BL��ȣ"				, 15
		ggoSpread.SSSetEdit 	C_BillSeqNo					, "����"				, 8

		ggoSpread.SSSetEdit 	C_CustItemCd				, "ǰ��"				, 15
		ggoSpread.SSSetEdit	    C_CustItemNm				, "ǰ���"				, 20
		ggoSpread.SSSetEdit		C_CustItemSpec				, "ǰ��԰�"			, 12
		ggoSpread.SSSetFloat		C_BillQty					, "����"				, 12,		parent.ggQtyNo,		  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_BillUnit					, "����"				, 12
		ggoSpread.SSSetFloat		C_BillPrc					, "�ܰ�"				, 15,		parent.ggUnitCostNo,  ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	 parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_BillDocAmt				, "�ݾ�"				, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetFloat		C_VatDocAmt					, "�ΰ����ݾ�"			, 15,		parent.ggAmtOfMoneyNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_VatType					, "�ΰ�������"			, 12
		ggoSpread.SSSetEdit		C_VatTypeNm					, "�ΰ���������"		, 20
		ggoSpread.SSSetFloat		C_VatRate					, "�ΰ�����"			, 12,		parent.ggExchRateNo,	ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,	,	,	"Z"
		ggoSpread.SSSetEdit		C_VatInc					, "�ΰ������Ա���"		, 15

		ggoSpread.SSSetEdit		C_ParentRowNo				, "C_ParentRowNo"		, 5
		ggoSpread.SSSetEdit		C_Flag						, "C_Flag"				, 5

'		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
'		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)


		Call ggoSpread.SSSetColHidden(C_BillNo, C_BillNo, True)
		Call ggoSpread.SSSetColHidden(C_BillSeqNo, C_BillSeqNo, True)
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
    ggoSpread.SpreadUnLock		C_Check , -1, -1
    ggoSpread.SpreadLock		C_BlNo,		-1,	C_TaxQty,		-1


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
	C_Check			    = 1			 '���� 
	C_CfmFlg		    = 2			 'Ȯ������ 
	C_BlNo			    = 3			 '����ó���ݰ�꼭��ȣ 
	C_BlCur			    = 4			 'ȭ�� 
	C_BlDocAmt		    = 5			 '���ް��� 
	C_BlVatDocAmt	    = 6			 '�ΰ����ݾ� 
	C_BlTotDocAmt	    = 7			 '�հ�ݾ� 
	C_BlVatType		    = 8			 '�ΰ������� 
	C_BlVatTypeNm	    = 9			 '�ΰ��������� 
	C_BlVatRt		    = 10		 '�ΰ����� 
	C_BlPayMeth		    = 11		 '������� 
	C_BlPayMethNm	    = 12		 '�������� 
	C_SoNo			    = 13		 '���ֹ�ȣ 
	C_CustPoNo		    = 14		 '���ֹ�ȣ 
	C_PoCompanyCd	    = 15		 '���ֹ��� 
	C_SoCompanyCd		= 16		 '���ֹ��� 
	C_TaxQty			= 17

End Sub

Sub InitSpreadPosVariables2()
	C_BillNo		= 1
	C_BillSeqNo		= 2

	C_CustItemCd	= 3         'ǰ�� 
	C_CustItemNm	= 4         'ǰ��� 
	C_CustItemSpec	= 5         'ǰ��԰� 
	C_BillQty		= 6         '���� 
	C_BillUnit		= 7         '���� 
	C_BillPrc		= 8         '�ܰ� 
	C_BillDocAmt	= 9         '�ݾ� 
	C_VatDocAmt		= 10        '�ΰ����ݾ� 
	C_VatType		= 11        '�ΰ������� 
	C_VatTypeNm		= 12        '�ΰ��������� 
	C_VatRate		= 13	    '�ΰ����� 
	C_VatInc		= 14        '�ΰ������Ա��� 

	C_ParentRowNo   = 15        '���� row ��ȣ 
	C_Flag          = 16        '�ڱ� ��ȣ 
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
				C_Check				=	iCurColumnPos(1)      'ȭ�� 
				C_CfmFlg		    =	iCurColumnPos(2)      '���ް��� 
				C_BlNo			    =	iCurColumnPos(3)      '�ΰ����ݾ� 
				C_BlCur			    =	iCurColumnPos(4)      '�հ�ݾ� 
				C_BlDocAmt		    =	iCurColumnPos(5)      '�ΰ������� 
				C_BlVatDocAmt	    =	iCurColumnPos(6)      '�ΰ��������� 
				C_BlTotDocAmt	    =	iCurColumnPos(7)      '�ΰ����� 
				C_BlVatType		    =	iCurColumnPos(8)      '������� 
				C_BlVatTypeNm	    =	iCurColumnPos(9)      '�������� 
				C_BlVatRt		    =	iCurColumnPos(10)     '���ֹ�ȣ 
				C_BlPayMeth		    =	iCurColumnPos(11)     '���ֹ�ȣ 
				C_BlPayMethNm	    =	iCurColumnPos(12)     '���ſ�û���¸� 
				C_SoNo			    =	iCurColumnPos(13)     '���ſ�û���� 
				C_CustPoNo		    =	iCurColumnPos(14)     '���ſ�û���и� 
				C_PoCompanyCd	    =	iCurColumnPos(15)     '���ֹ��� 
				C_SoCompanyCd		=	iCurColumnPos(16)     '���ֹ��� 
				C_TaxQty			=	iCurColumnPos(17)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_BillNo		=	iCurColumnPos(1)
				C_BillSeqNo		=	iCurColumnPos(2)
				C_CustItemCd	=	iCurColumnPos(3)        'ǰ�� 
				C_CustItemNm	=	iCurColumnPos(4)        'ǰ��� 
				C_CustItemSpec	=	iCurColumnPos(5)        'ǰ��԰� 
				C_BillQty		=	iCurColumnPos(6)        '���� 
				C_BillUnit		=	iCurColumnPos(7)        '���� 
				C_BillPrc		=	iCurColumnPos(8)        '�ܰ� 
				C_BillDocAmt	=	iCurColumnPos(9)        '�ݾ� 
				C_VatDocAmt		=	iCurColumnPos(10)        '�ΰ����ݾ� 
				C_VatType		=	iCurColumnPos(11)        '�ΰ������� 
				C_VatTypeNm		=   iCurColumnPos(12)       '�ΰ��������� 
				C_VatRate		=	iCurColumnPos(13)	    '�ΰ����� 
				C_VatInc		=	iCurColumnPos(14)       '�ΰ������Ա��� 

				C_ParentRowNo   =	iCurColumnPos(15)       '���� row ��ȣ 
				C_Flag          =	iCurColumnPos(16)       '�ڱ� ��ȣ 
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

	arrParam(0) = "���ֹ���"
	arrParam(1) = "B_BIZ_PARTNER"

	arrParam(2) = Trim(frm1.txtSoCompanyCd.Value)

	arrParam(4) = "BP_TYPE In ('S','CS') And IN_OUT_FLAG = 'O'"
	arrParam(5) = "���ֹ���"

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = "���ֹ���"
	arrHeader(1) = "���ֹ��θ�"

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

'------------------------------------------  OpenIvTeypCd()  -------------------------------------------------
' Name : OpenIvTeypCd()
' Description : SpreadItem PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenIvTeypCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "��������"
	arrParam(1) = "M_IV_TYPE"

	arrParam(2) = Trim(frm1.txtIvTypeCd.Value)

	arrParam(4) = "import_flg <> 'Y' and usage_flg = 'Y'"
	arrParam(5) = "��������"

	arrField(0) = "IV_TYPE_CD"
	arrField(1) = "IV_TYPE_NM"

	arrHeader(0) = "��������"
	arrHeader(1) = "�������¸�"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		frm1.txtIvTypeCd.focus
		Exit Function
	Else
		frm1.txtIvTypeCd.Value= arrRet(0)
		frm1.txtIvTypeNm.Value= arrRet(1)
		frm1.txtIvTypeCd.focus
	End If
End Function



'------------------------------------------  OpenPurGrp()  -------------------------------------------------
' Name : OpenPurGrp()
' Description : SpreadItem PopUp
'--------------------------------------------------------------------------------------------------------- ----
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "���ű׷�"
	arrParam(1) = "B_PUR_GRP"

	arrParam(2) = Trim(frm1.txtPurGrp.Value)

	arrParam(4) = ""
	arrParam(5) = "���ű׷�"

	arrField(0) = "PUR_GRP"
	arrField(1) = "PUR_GRP_NM"

	arrHeader(0) = "���ű׷�"
	arrHeader(1) = "���ű׷��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		frm1.txtPurGrp.focus
		Exit Function
	Else
		frm1.txtPurGrp.Value= arrRet(0)
		frm1.txtPurGrpNm.Value= arrRet(1)
		frm1.txtPurGrp.focus
	End If
End Function


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

'====================================== sprRedComColor() ======================================
'	Name : sprRedComColor()
'	Description : �������ڰ� ���� ���ں��� ������ ���� ��ȣ...
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
'	Description : �������ڰ� ���� ���ں��� ������ ���� ��ȣ...
'==============================================================================================
Sub sprBlackComColor(ByVal Col, ByVal Row, ByVal Row2)
    With frm1
		.vspdData2.Col = Col
		.vspdData2.Row = Row
        .vspdData2.ForeColor = &H0&
    End With
End Sub
'====================================== checkdt() ======================================
'	Name : checkdt()
'	Description : �������ڿ� ���� ����üũ.
'==============================================================================================
Sub checkdt(ByVal Row)
    With frm1
        .vspdData2.Row = Row
        .vspdData2.Col = C_PlanDt
        If UniConvDateToYYYYMMDD(.vspdData2.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(CurrDate,parent.gDateFormat,"") and Trim(.vspdData2.Text) <> "" Then
            Call sprRedComColor(C_PlanDt,Row,Row)
		else
		    Call sprBlackComColor(C_PlanDt,Row,Row)
        End If
    End With
End Sub



'==========================================   ApportionQtyChange()  ======================================
'	Name : ApportionQtyChange()
'	Description :
'=================================================================================================

Sub ApportionQtyChange(ByVal Row)
    Dim iparentrow
    Dim iReqQty,iApportionQty,iquotarate
    Dim totalquotarate,totalApportionQty
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim index

	with frm1.vspdData2
		.Row		= Row
		.Col		= C_ParentRowNo
		iparentrow  = Trim(.text)

		.Col		= C_Quota_Rate
		iquotarate  = Unicdbl(.text)

		lngRangeFrom = DataFirstRow(iparentrow)
	    lngRangeTo   = DataLastRow(iparentrow)

		totalquotarate = 0
		totalApportionQty = 0

		for index = lngRangeFrom  to lngRangeTo
		    .Row = index
		    .Col = 0
		    if Trim(.Text) <> ggoSpread.DeleteFlag  then
				.Col = C_Quota_Rate
				totalquotarate = totalquotarate + Unicdbl(.text)
		        if index <> clng(Row) then
				    .Col = C_ApportionQty
				    totalApportionQty = totalApportionQty + Unicdbl(.text)
		        End If
		    End If
		next

		frm1.vspdData.Row = iparentrow
		frm1.vspdData.Col = C_ReqQty
		iReqQty = Unicdbl(frm1.vspdData.text)

		'�հ� ������� 100�̸� ��η� = ��û�� - �����η��� 
		if totalquotarate = 100 then
		    iApportionQty = iReqQty - totalApportionQty
		else
			iApportionQty = (iquotarate * iReqQty)/100
	    End If

		.Row  = Row
		.Col  = C_ApportionQty
		.text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	End with
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
                Call DisplayMsgBox("17A005","X" ,"����ó", "X")
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
Function ShowFromData(Byval Row, Byval lngShowingRows)	'###�׸��� ������ ���Ǻκ�###
'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 3�� �����ϴ� ����� �����ϴ� �Լ���.
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

			'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� ù��° ���� 2��° ������ Row�� �����.
			If lngStartRow > 1 Then
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

			'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 7��° ���� ������ ������ Row�� �����.
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

			.Row = lngStartRow	'2003-03-01 Release �߰� 
			.Col = 0			'2003-03-01 Release �߰� 
			.Action = 0			'2003-03-01 Release �߰� 
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

        .Col = 1	'C_SupplierCd	'###�׸��� ������ ���Ǻκ�###
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
    	' �ش� �˻��׸�/������ ������ �ִ� ����ġ�� ���� 
    	Call DeleteDataForInsertSampleRows(lRow, lglngHiddenRows(lRow - 1))

    	' �� �߰� 
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

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()	'###�׸��� ������ ���Ǻκ�###

	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")  'Lock  Suitable  Field
	Call InitSpreadSheet
	Call InitSpreadSheet2
	Call InitVariables
	Call SetDefaultVal
    Call CookiePage(0)
	set gActiveSpdSheet = frm1.vspdData
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
		If DbQuery2(lgCurrRow, False) = False Then	Exit Sub
	End If
	frm1.vspdData.redraw = true
End Sub



'======================================================================================================
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'=======================================================================================================
'==========================================================================================
'   Event Name : txtFrBillDt
'   Event Desc :
'==========================================================================================
Sub txtFrBillDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrBillDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtFrBillDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtToBillDt
'   Event Desc :
'==========================================================================================
Sub txtToBillDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToBillDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtToBillDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtBillDt
'   Event Desc :
'==========================================================================================
Sub txtBillDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBillDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtBillDt.Focus
	End if
End Sub
Sub txtPayDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPayDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtPayDt.Focus
	End if
End Sub


'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtFrBillDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToBillDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtBillDt_KeyDown(KeyCode, Shift)
'	If KeyCode = 13	Then Call MainQuery()
End Sub


'======================================================================================================
' Function Name : FncQuery
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

    If Not chkField(Document, "1") Then
       		Exit Function
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
'    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If


' === 2005.07.22 ���� ===========================================================

	If ValidDateCheck(frm1.txtFrbillDt, frm1.txtToBillDt) = False Then Exit Function

' === 2005.07.22 ���� ===========================================================

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

    Dim IntRetCD

	'-----------------------
    'Precheck area
    '-----------------------
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ START
'	If DefaultCheck = False Then
'		Exit Function
'	End If
    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ END

'	intRetCd = DisplayMsgBox("900018", VB_YES_NO, "X", "X")   '�� �ٲ�Eκ?
'	If intRetCd = VBNO Then
'		Exit Function
'	End IF


    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then
       		Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If


    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck  Then
       Exit Function
    End If

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
    Call parent.FncFind(Parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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

	with frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then

		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtFrBillDt=" & Trim(.txtFrBillDt.Text)
			strVal = strVal & "&txtToBillDt=" & Trim(.txtToBillDt.Text)
			strVal = strVal & "&txtBillDt=" & Trim(.txtBillDt.Text)
			strVal = strVal & "&txtCustPoNo=" & Trim(.txtCustPoNo.value)
			strVal = strVal & "&txtBlNo=" & Trim(.txtBlNo.value)
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '��: Next key tag
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtFrBillDt=" & Trim(.txtFrBillDt.Text)
			strVal = strVal & "&txtToBillDt=" & Trim(.txtToBillDt.Text)
			strVal = strVal & "&txtBillDt=" & Trim(.txtBillDt.Text)
			strVal = strVal & "&txtCustPoNo=" & Trim(.txtCustPoNo.value)
			strVal = strVal & "&txtBlNo=" & Trim(.txtBlNo.value)
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '��: Next key tag
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
	Call SetToolBar("11001000000011")				'��ư ���� ���� 

	frm1.btnSelect.disabled = false
	frm1.btnDisSelect.disabled = false
	frm1.btnConfirm.disabled = false
	frm1.btnCancel.disabled = false

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


	'/* 9�� ������ġ: ���� ���������� �ణ �̵� �� �̹� ��ȸ�� �ڷᳪ �Էµ� �ڷḦ �о� ���� ������ '' â ���� - START */
	Call LayerShowHide(1)

	With frm1
		.vspdData.redraw = false
		.vspdData.Row = CInt(Row)
		.vspdData.Col = .vspdData.MaxCols
		pRow  = CInt(.vspdData.Text)

		If lglngHiddenRows(pRow - 1) <> 0 And NextQueryFlag = False Then
			.vspdData2.ReDraw = False
			lngRet = ShowFromData(pRow, lglngHiddenRows(pRow - 1))	'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 3�� �����ϴ� ����� �����ϴ� �Լ���.
			Call SetToolBar("11001111001011")				'��ư ���� ���� 
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
		.vspdData.Col = C_PoCompanyCd
		strVal = strVal & "&txtPoCompanyCd=" & trim(.vspdData.text)
		.vspdData.Row = Row
		.vspdData.Col = C_SoCompanyCd
		strVal = strVal & "&txtSoCompanyCd=" & trim(.vspdData.text)
		.vspdData.Row = Row
		.vspdData.Col = C_BlNo
		strVal = strVal & "&txtBlNo=" & trim(.vspdData.text)
		strVal = strVal & "&lgStrPrevKeyM="  & lgStrPrevKeyM(Row - 1)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'��: Next key tag
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(pRow)

'		msgbox strVal
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

	Dim lRow
	Dim lGrpCnt, i, j, iCnt
	Dim strVal, strDel, strTxt, k

	DbSave = False

	Call LayerShowHide(1)
	On Error Resume Next                                                   <%'��: Protect system from crashing%>

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

		    Case ggoSpread.InsertFlag											'��: �ű� 

					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep				'��: C=Create
		                        strVal = strVal & parent.gRowSep

		        		lGrpCnt = lGrpCnt + 1

		    Case ggoSpread.UpdateFlag											'��: �ű� 
		    	.vspdData.Col = C_Check
				If Trim(.vspdData.Text) = "1" then
					.vspdData.Col = C_TaxQty
					If Trim(UNICDbl(.vspdData.Text)) = "" Or Trim(UNICDbl(.vspdData.Text)) < 0 Then
						Call DisplayMsgBox("175220","X","X","X")
						Call LayerShowHide(0)
						.vspdData.Col = 0
						.vspdData.Text = ""
					Else
							strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep				'��: U=Update
							'��� �������� 
							strVal = strVal & Trim(.txtIvTypeCd.value) & parent.gColSep
							strVal = strVal & UNIConvDate(Trim(.txtBillDt.text)) & parent.gColSep
							strVal = strVal & Trim(.txtPurGrp.value) & parent.gColSep
							.vspdData.Col = C_Check		 								'���ÿ��� 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_CfmFlg		 							'Ȯ������ 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_BlNo		 								'BL��ȣ 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_BlCur		 								'ȭ�� 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_BlDocAmt		 								'���ް��� 
							strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
							.vspdData.Col = C_BlVatDocAmt		 								'�ΰ����ݾ� 
							strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
							.vspdData.Col = C_BlTotDocAmt		 								'�հ�ݾ� 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_BlVatType		 								'�ΰ������� 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_BlVatRt		 								'�ΰ����� 
							strVal = strVal & UNICDbl(Trim(.vspdData.Text)) & parent.gColSep
							.vspdData.Col = C_BlPayMeth		 								'������� 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_SoNo		 								'���ֹ�ȣ 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_CustPoNo		 								'���ֹ�ȣ 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_PoCompanyCd		 								'���ֹ�ȣ 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							.vspdData.Col = C_SoCompanyCd		 								'���ֹ�ȣ 
							strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
							strVal = strVal & UNIConvDate(Trim(.txtPayDt.text)) & parent.gColSep		'������ 
							strVal = strVal & parent.gRowSep

		        			lGrpCnt = lGrpCnt + 1
					End If

				End If

		    Case ggoSpread.DeleteFlag													'��: ���� 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep						'��: U=Update
		                        strVal = strVal & parent.gRowSep

				        lGrpCnt = lGrpCnt + 1
		End Select

    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'��: �����Ͻ� ASP �� ���� %>

	End With

    DbSave = True
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData2.MaxRows = 0

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If


	'-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then
		Exit Function
	End If

End Function


'==========================================================================================
'   Event Name : btnSelect_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i, ClsFlg

	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Row = i

			'If ClsFlg <> "Y" Then
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_Check, i, 1)
			'End If
		Next

	End If
End Sub
'==========================================================================================
'   Event Name : btnDisSelect_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i,ClsFlg_1

	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Row = i

			frm1.vspdData.Col = C_Cfmflg
			ClsFlg_1 = frm1.vspdData.text

			If ClsFlg_1 = "1" Then
				frm1.vspdData.Col = C_Cfmflg
				frm1.vspdData.value = 0
				Call vspdData_ButtonClicked(C_Cfmflg, i, 0)
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 0
				Call vspdData_ButtonClicked(C_Check, i, 0)
			else
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 0
				Call vspdData_ButtonClicked(C_Check, i, 0)
			End If
		Next
	End If
End Sub
'==========================================================================================
'   Event Name : btnConfirm_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnConfirm_OnClick()
	Dim i, ClsFlg_2

	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Row = i

			frm1.vspdData.Col = C_Check
			ClsFlg_2 = frm1.vspdData.text

			If ClsFlg_2 = "0" Then
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_Check, i, 1)
				frm1.vspdData.Col = C_CfmFlg
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_CfmFlg, i, 1)

			Else
				frm1.vspdData.Col = C_CfmFlg
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_CfmFlg, i, 1)
			End If
		Next

	End If
End Sub
'==========================================================================================
'   Event Name : btnCancel_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnCancel_OnClick()
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

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Check And Row > 0 Then
		ggoSpread.Source = frm1.vspdData
	    Select Case ButtonDown
	    Case 1
			ggoSpread.UpdateRow Row
	    Case 0
			frm1.vspdData.Col = C_CfmFlg
			if frm1.vspdData.text = 1 then
				frm1.vspdData.value = 0
			end if
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row
			frm1.vspdData.text = ""
			lgBlnFlgChgValue = False
	    End Select
	ElseIf Col = C_CfmFlg And Row > 0 Then
		ggoSpread.Source = frm1.vspdData
	    Select Case ButtonDown
	    Case 1
			frm1.vspdData.Col = C_Check
			if frm1.vspdData.text = 0 then
				frm1.vspdData.value = 1
			end if
			ggoSpread.UpdateRow Row
	    Case 0
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row
'			frm1.vspdData.text = ""
'			lgBlnFlgChgValue = False
	    End Select
	End If
	lgBlnFlgChgValue = True

End Sub


'--------------------------------------------------------------------
'		Cookie ����Լ� 
'--------------------------------------------------------------------
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

		With frm1
			'���ֹ���/����ó��꼭������/���ֹ�ȣ/����ó��꼭��ȣ 
			WriteCookie "txtSoCompanyCd", Trim(.txtSoCompanyCd.value)
			WriteCookie "txtFrBillDt", Trim(.txtFrBillDt.value)
			WriteCookie "txtToBillDt", Trim(.txtToBillDt.value)
			WriteCookie "txtCustPoNo", Trim(.txtCustPoNo.value)
			WriteCookie "txtBlNo", Trim(.txtBlNo.value)
		End With

		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)
	End IF
End Function

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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��Ƽ���۴� ����</font></TD>
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
								<TD CLASS="TD5" NOWRAP>���ֹ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="���ֹ���"   NAME="txtSoCompanyCd" SIZE=10 MAXLENGTH=10 tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenSoCompany" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSoCompany()" >
														<INPUT TYPE=TEXT ALT="���ֹ���" NAME="txtSoCompanyNm" SIZE=20 MAXLENGTH=50 tag="14X">
								<TD CLASS="TD5" NOWRAP>����ó��꼭������</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm111ma1_fpDateTime2_txtFrBillDt.js'></script>~
										<script language =javascript src='./js/kmm111ma1_fpDateTime2_txtToBillDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="���ֹ�ȣ"   NAME="txtCustPoNo" SIZE=33 MAXLENGTH=18 tag="11NXXU" >
								<!--<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()">-->
								</TD>
								<TD CLASS="TD5" NOWRAP>����ó��꼭��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����ó��꼭��ȣ"   NAME="txtBlNo" SIZE=35 MAXLENGTH=18 tag="11NXXU" ></TD>
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
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="��������"   NAME="txtIvTypeCd" SIZE=10 MAXLENGTH=5 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvTypeCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvTeypCd()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
								     <INPUT TYPE=TEXT ALT="��������" NAME="txtIvTypeNm" SIZE=20 MAXLENGTH=50 tag="24X">
								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm111ma1_fpDateTime1_txtBillDt.js'></script>
<!--										<script language =javascript src='./js/kmm111ma1_fpDateTime2_txtBillDt.js'></script>		-->
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
								<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT ALT="���ű׷�"   NAME="txtPurGrp" SIZE=10 MAXLENGTH=4 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
									<INPUT TYPE=TEXT ALT="���ű׷�" NAME="txtPurGrpNm" SIZE=20 MAXLENGTH=50 tag="24X">
								<TD CLASS="TD5" NOWRAP>������</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm111ma1_fpDateTime1_txtPayDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=45% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm111ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=45% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm111ma1_B_vspdData2.js'></script>
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
					<TD WIDTH="*" align="left">
						<BUTTON NAME="btnSelect" CLASS="CLSMBTN">�ϰ�����</BUTTON>
						<BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">�ϰ��������</BUTTON>
						<BUTTON NAME="btnConfirm" CLASS="CLSMBTN">�ϰ�Ȯ��</BUTTON>
						<BUTTON NAME="btnCancel" CLASS="CLSMBTN">�ϰ�Ȯ�����</BUTTON>
					</TD>
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">��Ƽ���۴ϸ�����ȸ</a></td>
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
<INPUT TYPE=HIDDEN NAME="hdnUsrId" tag="24">
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