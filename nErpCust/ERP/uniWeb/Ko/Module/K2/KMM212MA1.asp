<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : MM212MA1
'*  4. Program Name         : ��Ƽ���۴�B/LȮ��/����-��Ƽ 
'*  5. Program Desc         : ��Ƽ���۴�B/LȮ��/����-��Ƽ 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2005/03/07
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "KMM212MB1.asp"
Const BIZ_PGM_ID2 = "KMM212MB101.asp"
'Const BIZ_PGM_SAVE_ID = "m2111mb5.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL = "KMM211MA1"
											'��: �����Ͻ� ���� ASP�� 
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'��� �������� 
'==========================================================
Dim C_Check

Dim	C_BpCd						'���ֹ��� 
Dim C_BpNm                      '���ֹ��θ� 
Dim C_BlNo		                'B/L��ȣ 
Dim C_BlDocNo                   'B/L������ȣ 
Dim C_PostedFlg                 'B/LȮ������ 
Dim C_BlIssueDt                 'B/L������ 
Dim C_LoadingDt                 '������ 
Dim C_Currency                  'ȭ�� 
Dim C_DocAmt                    'B/L�ݾ� 
Dim C_LocAmt                    'B/L�ڱ��ݾ� 
Dim C_XchRate	                'ȯ�� 
Dim C_IvType	                '�������� 
Dim C_IvTypeNm                  '�������¸� 
Dim C_PayMethod                 '������� 
Dim C_PayMethodNm               '��������� 
Dim C_PurGrp                    '���ű׷� 
Dim C_PurGrpNm			        '���ű׷�� 
Dim C_Beneficiary
Dim C_Applicant
Dim C_PostedYN					'Ȯ������ 
Dim C_RefIvNo



'�ϴ� �������� 
'==========================================================
Dim C_DBlNo						'B/L��ȣ 
Dim C_BlSeq						'B/L�Ϸù�ȣ 

Dim C_ItemCd					'ǰ�� 
Dim C_ItemNm					'ǰ��� 
Dim C_Spec						'ǰ��԰� 
Dim C_Qty						'���� 
Dim C_Unit						'���� 
Dim C_Price						'�ܰ� 
Dim C_DDocAmt					'�ݾ� 
Dim C_DLocAmt					'�ΰ����ݾ� 
Dim C_PoNo						'�ΰ������� 
Dim C_PoSeq						'�ΰ��������� 

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

	frm1.txtBlIssueFrDt.Text = StartDate
	frm1.txtBlIssueToDt.Text = EndDate
'	frm1.txtBlIssueDt.Text = EndDate

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

    .MaxCols = C_RefIvNo + 1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")
		ggoSpread.SSSetCheck	C_Check								  	,	"����"		,10,,,true

		ggoSpread.SSSetEdit		C_BpCd								  	,	"���ֹ���"		,10
		ggoSpread.SSSetEdit		C_BpNm                                	,	"���ֹ��θ�"	,18
		ggoSpread.SSSetEdit		C_BlNo		                          	,	"B/L��ȣ"       ,18
		ggoSpread.SSSetEdit		C_BlDocNo                             	,	"B/L������ȣ"   ,18
		ggoSpread.SSSetEdit		C_PostedFlg                           	,	"B/LȮ������"   ,15
		ggoSpread.SSSetEdit		C_BlIssueDt                           	,	"B/L������"     ,18
		ggoSpread.SSSetEdit		C_LoadingDt                           	,	"������"        ,18
		ggoSpread.SSSetEdit		C_Currency                            	,	"ȭ��"          ,10
		SetSpreadFloatLocal		C_DocAmt                              	,	"B/L�ݾ�"       ,15,1,5
		SetSpreadFloatLocal		C_LocAmt                              	,	"B/L�ڱ��ݾ�"   ,15,1,5
		SetSpreadFloatLocal		C_XchRate	                          	,	"ȯ��"          ,12,1,5
		ggoSpread.SSSetEdit		C_IvType	                          	,	"��������"      ,10
		ggoSpread.SSSetEdit		C_IvTypeNm                            	,	"�������¸�"    ,18
		ggoSpread.SSSetEdit		C_PayMethod                           	,   "�������"      ,10
		ggoSpread.SSSetEdit		C_PayMethodNm                         	,   "���������"    ,18
		ggoSpread.SSSetEdit		C_PurGrp                              	,   "���ű׷�"      ,10
		ggoSpread.SSSetEdit		C_PurGrpNm		                      	,   "���ű׷��"    ,18
		ggoSpread.SSSetEdit		C_Beneficiary							,	"������"		,18
		ggoSpread.SSSetEdit		C_Applicant								,	"������"		,18
		ggoSpread.SSSetEdit		C_PostedYN								,	"Ȯ������"		,18
		ggoSpread.SSSetEdit		C_RefIvNo								,	"���Թ�ȣ"		,18



'	Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)
	Call ggoSpread.SSSetColHidden(C_BpCd,	C_BpCd,	True)
	Call ggoSpread.SSSetColHidden(C_BpNm,	C_BpNm,	True)
	Call ggoSpread.SSSetColHidden(C_Beneficiary,	C_Beneficiary,	True)
	Call ggoSpread.SSSetColHidden(C_Applicant,	C_Applicant,	True)
	Call ggoSpread.SSSetColHidden(C_PostedYN,	C_PostedYN,	True)
	Call ggoSpread.SSSetColHidden(C_RefIvNo,	C_RefIvNo,	True)

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


		ggoSpread.SSSetEdit 	C_DBlNo						, "BL��ȣ"				, 15
		ggoSpread.SSSetEdit 	C_BlSeq						, "����"				, 5

		ggoSpread.SSSetEdit 	C_ItemCd					, "ǰ��"				, 15
		ggoSpread.SSSetEdit	    C_ItemNm					, "ǰ���"				, 20
		ggoSpread.SSSetEdit		C_Spec						, "ǰ��԰�"			, 20
		SetSpreadFloatLocal		C_Qty						, "����"				, 12,1,5
		ggoSpread.SSSetEdit		C_Unit						, "����"				, 10
		SetSpreadFloatLocal		C_Price						, "�ܰ�"				, 15,1,5
		SetSpreadFloatLocal		C_DDocAmt					, "�ݾ�"				, 15,1,5
		SetSpreadFloatLocal		C_DLocAmt					, "�ڱ��ݾ�"			, 15,1,5
		ggoSpread.SSSetEdit		C_PoNo						, "���ֹ�ȣ"			, 15
		ggoSpread.SSSetEdit		C_PoSeq						, "���ּ���"			, 10

		ggoSpread.SSSetEdit		C_ParentRowNo				, "C_ParentRowNo"		, 5
		ggoSpread.SSSetEdit		C_Flag						, "C_Flag"				, 5

'		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
'		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)

		Call ggoSpread.SSSetColHidden(C_DBlNo,C_DBlNo, True)
		Call ggoSpread.SSSetColHidden(C_BlSeq,C_BlSeq, True)

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
    ggoSpread.SpreadLock		C_BpCd,		-1,	C_RefIvNo,		-1

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
'    With frm1

'    .vspdData.ReDraw = False
'    ggoSpread.SSSetProtected		C_PlantCd, pvStartRow, pvEndRow


'    .vspdData.ReDraw = True

'    End With
End Sub

Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
'    With frm1

'    .vspdData2.ReDraw = False

'	ggoSpread.SSSetRequired  C_SpplCd,		    pvStartRow,	pvEndRow
'	ggoSpread.SSSetProtected C_SpplNm,		    pvStartRow,	pvEndRow
'	ggoSpread.SSSetRequired  C_Quota_Rate,		pvStartRow,	pvEndRow
'    ggoSpread.SSSetRequired  C_ApportionQty,	pvStartRow,	pvEndRow
'    ggoSpread.SSSetRequired  C_PlanDt,			pvStartRow,	pvEndRow
'    ggoSpread.SSSetRequired  C_GrpCd,			pvStartRow,	pvEndRow
'    ggoSpread.SSSetProtected C_GrpNm,		    pvStartRow,	pvEndRow

'   .vspdData2.ReDraw = True
'    End With
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_Check					=	1		  '���� 
	C_BpCd					=	2   	  '���ֹ��� 
	C_BpNm                  =	3         '���ֹ��θ� 
	C_BlNo		            =	4         'B/L��ȣ 
	C_BlDocNo               =	5         'B/L������ȣ 
	C_PostedFlg             =	6         'B/LȮ������ 
	C_BlIssueDt             =	7         'B/L������ 
	C_LoadingDt             =	8         '������ 
	C_Currency              =	9         'ȭ�� 
	C_DocAmt                =	10        'B/L�ݾ� 
	C_LocAmt                =	11        'B/L�ڱ��ݾ� 
	C_XchRate	            =	12        'ȯ�� 
	C_IvType	            =	13        '�������� 
	C_IvTypeNm              =	14        '�������¸� 
	C_PayMethod             =	15        '������� 
	C_PayMethodNm           =	16        '��������� 
	C_PurGrp                =	17        '���ű׷� 
	C_PurGrpNm		        =	18	      '���ű׷�� 
	C_Beneficiary			=	19
	C_Applicant	 			=	20
	C_PostedYN				=	21
	C_RefIvNo				=	22

End Sub

Sub InitSpreadPosVariables2()
	C_DBlNo					=	1		'B/L��ȣ 
	C_BlSeq		            =	2       'B/L�Ϸù�ȣ 

	C_ItemCd	            =	3       'ǰ�� 
	C_ItemNm	            =	4       'ǰ��� 
	C_Spec		            =	5       'ǰ��԰� 
	C_Qty		            =	6       '���� 
	C_Unit		            =	7       '���� 
	C_Price		            =	8       '�ܰ� 
	C_DDocAmt	            =	9       '�ݾ� 
	C_DLocAmt	            =	10      '�ڱ��ݾ� 
	C_PoNo		            =	11      '���ֹ�ȣ 
	C_PoSeq		            =	12      '���ּ��� 


	C_ParentRowNo			=	13       '���� row ��ȣ 
	C_Flag					=	14       '�ڱ� ��ȣ 
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
				C_Check					=	iCurColumnPos(1)
				C_BpCd					=	iCurColumnPos(2)        '���ֹ��� 
				C_BpNm                  =	iCurColumnPos(3)        '���ֹ��θ� 
				C_BlNo		            =	iCurColumnPos(4)        'B/L��ȣ 
				C_BlDocNo               =	iCurColumnPos(5)        'B/L������ȣ 
				C_PostedFlg             =	iCurColumnPos(6)        'B/LȮ������ 
				C_BlIssueDt             =	iCurColumnPos(7)        'B/L������ 
				C_LoadingDt             =	iCurColumnPos(8)        '������ 
				C_Currency              =	iCurColumnPos(9)        'ȭ�� 
				C_DocAmt                =	iCurColumnPos(10)       'B/L�ݾ� 
				C_LocAmt                =	iCurColumnPos(11)       'B/L�ڱ��ݾ� 
				C_XchRate	            =	iCurColumnPos(12)       'ȯ�� 
				C_IvType	            =	iCurColumnPos(13)       '�������� 
				C_IvTypeNm              =	iCurColumnPos(14)       '�������¸� 
				C_PayMethod             =	iCurColumnPos(15)       '������� 
				C_PayMethodNm           =	iCurColumnPos(16)       '��������� 
				C_PurGrp                =	iCurColumnPos(17)       '���ű׷� 
				C_PurGrpNm		        =	iCurColumnPos(18)		'���ű׷�� 
				C_Beneficiary	        =	iCurColumnPos(19)		'
				C_Applicant		        =	iCurColumnPos(20)		'
				C_PostedYN				=	iCurColumnPos(21)
				C_RefIvNo				=	iCurColumnPos(22)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_DBlNo					=	iCurColumnPos(1)		'B/L��ȣ 
				C_BlSeq		            =	iCurColumnPos(2)       'B/L�Ϸù�ȣ 
				C_ItemCd	            =	iCurColumnPos(3)       'ǰ�� 
				C_ItemNm	            =	iCurColumnPos(4)       'ǰ��� 
				C_Spec		            =	iCurColumnPos(5)       'ǰ��԰� 
				C_Qty		            =	iCurColumnPos(6)       '���� 
				C_Unit		            =	iCurColumnPos(7)       '���� 
				C_Price		            =	iCurColumnPos(8)       '�ܰ� 
				C_DDocAmt	            =	iCurColumnPos(9)       '�ݾ�             
				C_DLocAmt	            =	iCurColumnPos(10)      '�ڱ��ݾ�         
				C_PoNo		            =	iCurColumnPos(11)      '���ֹ�ȣ         
				C_PoSeq		            =	iCurColumnPos(12)      '���ּ���         
				C_ParentRowNo           =	iCurColumnPos(13)  
				C_Flag                  =	iCurColumnPos(14)  
	End Select
End Sub	

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
'    Call CookiePage(0)
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
'   Event Name : txtBlIssueFrDt
'   Event Desc :
'==========================================================================================
Sub txtBlIssueFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBlIssueFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtBlIssueFrDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtBillToDt
'   Event Desc :
'==========================================================================================
Sub txtBlIssueToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBlIssueToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtBlIssueToDt.Focus
	End if
End Sub

Sub txtBlIssueDt_DblClick(Button)
	if Button = 1 then
		frm1.txtBlIssueDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtBlIssueDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtBlIssueFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtBlIssueToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtBlIssueDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
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
    If Not chkField(Document, "1") Then											'This function check indispensable field
	   Exit Function
    End If

' === 2005.07.22 ���� ===========================================================

	If ValidDateCheck(frm1.txtBlIssueFrDt, frm1.txtBlIssueToDt) = False Then Exit Function

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
	If DefaultCheck = False Then
		Exit Function
	End If
    '8�� ������ġ: ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ END

'	intRetCd = DisplayMsgBox("900018", VB_YES_NO, "X", "X")   '�� �ٲ�Eκ?
'	If intRetCd = VBNO Then
'		Exit Function
'	End IF

	If frm1.rdoCfmflg2.checked Then
    '-----------------------
    'Check content area
    '-----------------------
		If Not chkField(Document, "2") Then
		   		Exit Function
		End If
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

<!--
'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
-->
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

    	For lDelRow = .SelBlockRow to .SelBlockRow2
    		Call deleteSum(lDelRow)
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_Check, lDelRow, 1)
    	Next

    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------

    lgBlnFlgChgValue = True
    'Call TotalSum

	'------ Developer Coding part (End )   --------------------------------------------------------------
    If Err.number = 0 Then
       FncDeleteRow = True                                                            '��: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement

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
			strVal = strVal & "&txtBlIssueFrDt=" & Trim(.txtBlIssueFrDt.Text)
			strVal = strVal & "&txtBlIssueToDt=" & Trim(.txtBlIssueToDt.Text)
			strVal = strVal & "&rdoCfmflg=" & Trim(.rdoCfmflg.Text)
			strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
			strVal = strVal & "&txtBlNo=" & Trim(.txtBlNo.value)
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '��: Next key tag
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else

		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtBlIssueFrDt=" & Trim(.txtBlIssueFrDt.Text)
			strVal = strVal & "&txtBlIssueToDt=" & Trim(.txtBlIssueToDt.Text)
			if .rdoCfmflg(0).checked = true then
				strVal = strVal & "&rdoCfmflg=" & "Y"
			else
				strVal = strVal & "&rdoCfmflg=" & "N"
			End if
			strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
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


	If frm1.rdoCfmflg3.checked Then
		Call SetToolBar("11001010000011")				'��ư ���� ���� 
	Else
		Call SetToolBar("11001000000011")				'��ư ���� ���� 
	End If


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
		.vspdData.Col = C_BlDocNo
		strVal = strVal & "&txtBlNo=" & trim(.vspdData.text)
		strVal = strVal & "&lgStrPrevKeyM="  & lgStrPrevKeyM(Row - 1)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'��: Next key tag
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(pRow)

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
'======================================  DbSave()  =================================
Function DbSave()
    Dim lRow, DRow
    Dim lGrpCnt
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim iOrderQty
	Dim iCost
	Dim iOrderAmt

	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�]
	Dim strDTotalvalLen  '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]

	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�]
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size

	Dim iTmpDBuffer          '������ ���� [����]
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size
    Dim ii


	DbSave = False

	ColSep = Parent.gColSep
	RowSep = Parent.gRowSep

	With frm1
'		.txtMode.value = Parent.UID_M0002

		lGrpCnt = 0

		strVal = ""
		strDel = ""
		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0


    For lRow = 1 To .vspdData.MaxRows step 1
        Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		     Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
	     		.txtMode.value = Parent.UID_M0002
				if Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))=ggoSpread.InsertFlag then
					strVal = "C" & ColSep
				Else
					strVal = "U" & ColSep
				End if

				strVal = strVal & Trim(GetSpreadText(.vspdData,C_BlDocNo,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(frm1.txtBlIssueDt.Text) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PostedYn,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_RefIvNo,lRow,"X","X")) & ColSep
                strVal = strVal & lRow & RowSep

				lGrpCnt = lGrpCnt + 1

            Case ggoSpread.DeleteFlag
				.txtMode.value = Parent.UID_M0003
'				If Trim(UNICDbl(GetSpreadText(.vspdData,C_PostedFlg,lRow,"X","X"))) = "Ȯ��" then
'					Call DisplayMsgBox("970021", "X","B/LȮ���� ����ϼ���", "X")
'					.vspdData.Row = lRow
'					.vspdData.Action = 0
'					Call LayerShowHide(0)
'					Exit Function
'				End if

				strDel = strDel & "D" & ColSep
				.vspdData.Col = C_BlDocNo
				strDel = strDel & Trim(.vspdData.Text) & ColSep
				.vspdData.Col = C_Beneficiary
				strDel = strDel & Trim(.vspdData.Text) & ColSep
				.vspdData.Col = C_Applicant
				strDel = strDel & Trim(.vspdData.Text) & ColSep
				.vspdData.Col = C_PurGrp
				strDel = strDel & Trim(.vspdData.Text) & ColSep
                strDel = strDel & lRow & RowSep

				lGrpCnt = lGrpCnt + 1

        End Select

		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 

		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)

		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If

		         iTmpCUBufferCount = iTmpCUBufferCount + 1

		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		   Case ggoSpread.DeleteFlag

		         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)

		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0
		         End If

		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If

		         iTmpDBuffer(iTmpDBufferCount) =  strDel
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
		End Select
	Next
	frm1.txtSpread.value  = strDel
	frm1.txtMaxRows.value = lGrpCnt
    End With

	frm1.txtMaxRows.value = lGrpCnt
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If

	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If

	If LayerShowHide(1) = False Then Exit Function
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
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

' 	with frm1
'		if (UniConvDateToYYYYMMDD(.txtReqFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtReqToDt.text,Parent.gDateFormat,"")) and trim(.txtReqFrDt.text)<>"" and trim(.txtReqToDt.text)<>"" then
'			Call DisplayMsgBox("17a003", "X","��û��", "X")
'			Exit Function
'		End If

'		if (UniConvDateToYYYYMMDD(.txtDlvyFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtDlvyToDt.text,Parent.gDateFormat,"")) and trim(.txtDlvyFrDt.text)<>"" and trim(.txtDlvyToDt.text)<>"" then
'			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")
'			Exit Function
'		End If

'	End with

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
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Check
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 0 then
			    frm1.vspdData.value = 1
                Call vspdData_ButtonClicked(C_Check, i, 1)
		    end if
		Next
	End If
End Sub
'==========================================================================================
'   Event Name : btnDisSelect_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Check
			frm1.vspdData.Row = i
			if frm1.vspdData.value = 1 then
			    frm1.vspdData.value = 0
                Call vspdData_ButtonClicked(C_Check, i, 0)
		    end if
		Next
	End If
End Sub

'==========================================================================================
'   Event Name : btnConfirm_OnClick()
'   Event Desc :
'==========================================================================================
Sub btnConfirm_OnClick()
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
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row
			frm1.vspdData.text = ""
			lgBlnFlgChgValue = False
	    End Select
	ElseIf Col = C_CfmFlg And Row > 0 Then
		ggoSpread.Source = frm1.vspdData
	    Select Case ButtonDown
	    Case 1
			ggoSpread.UpdateRow Row
	    Case 0
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row
			frm1.vspdData.text = ""
			lgBlnFlgChgValue = False
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

'	    If lgIntFlgMode <> Parent.OPMD_UMODE Then
'	        Call DisplayMsgBox("900002", "X", "X", "X")
'	        Exit Function
'	    End If

'	    If lgBlnFlgChgValue = True Then
'			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
'			If IntRetCD = vbNo Then
'				Exit Function
'			End If
'	    End If

'		WriteCookie "PoNo" , frm1.txtPoNo.value

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
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��Ƽ���۴�BL Ȯ��/����</font></TD>
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
														<INPUT TYPE=TEXT ALT="���ֹ���" NAME="txtSoCompanyNm" SIZE=20 MAXLENGTH=50 tag="24X">
								<TD CLASS="TD5" NOWRAP>B/L������</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm212ma1_fpDateTime2_txtBlIssueFrDt.js'></script>~
										<script language =javascript src='./js/kmm212ma1_fpDateTime2_txtBlIssueToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>B/LȮ��ó������</TD>
								<TD CLASS="TD6" NOWRAP><!--<INPUT TYPE=radio Class="Radio" ALT="B/LȮ��ó������" NAME="rdoCfmflg" id = "rdoCfmflg1" Value="A" checked tag="11"><label for="rdoCfmflg1">&nbsp;��ü&nbsp;</label> -->
													   <INPUT TYPE=radio Class="Radio" ALT="B/LȮ��ó������" NAME="rdoCfmflg" id = "rdoCfmflg2" Value="Y" tag="11"><label for="rdoCfmflg2">&nbsp;Ȯ��&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="B/LȮ��ó������" NAME="rdoCfmflg" id = "rdoCfmflg3" Value="N" tag="11" checked><label for="rdoCfmflg3">&nbsp;��Ȯ��&nbsp;</label></TD>
								<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="���ֹ�ȣ"   NAME="txtPoNo" SIZE=35 MAXLENGTH=18 tag="11NXXU" ></TD>

							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>B/L��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="B/L��ȣ"   NAME="txtBlNo" SIZE=35 MAXLENGTH=18 tag="11NXXU" ></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>&nbsp;</TD>

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
								<TD HEIGHT=10 WIDTH=100% CLASS="TD5" NOWRAP>B/L������</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/kmm212ma1_fpDateTime2_txtBlIssueDt.js'></script></TD>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
							</TR>
							<TR>
								<TD HEIGHT=70% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm212ma1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=* WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm212ma1_B_vspdData2.js'></script>
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
					</TD>
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">��Ƽ���۴�B/L���</a></td>
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
<INPUT TYPE=HIDDEN NAME="txtPostedFlg" tag="24">
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