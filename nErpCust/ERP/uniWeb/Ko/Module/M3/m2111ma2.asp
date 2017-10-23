<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : m2111ma2
'*  4. Program Name         : ��ü���� 
'*  5. Program Desc         : ��ü���� 
'*  6. Component List       :
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
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

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
CONST BIZ_PGM_ID = "M2111MB2.ASP"
CONST BIZ_PGM_ID2 = "m2111mb201.asp"
CONST BIZ_PGM_SAVE_ID = "M2111MB2.ASP"

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_CfmFlg      '���� 
Dim C_ReqNo       '��û��ȣ 
Dim C_PlantCd     '���� 
Dim C_PlantNm     '����� 
Dim C_ItemCd      'ǰ�� 
Dim C_ItemNm      'ǰ��� 
Dim C_SpplSpec    '�԰� 
Dim C_ReqQty      '��û�� 
Dim C_Unit        '���� 
Dim C_TrackingNo
Dim C_ReqDt       '�ʿ��� 
Dim C_ReqStateCd  '���ſ�û���� 
Dim C_ReqStateNm  '���ſ�û���¸� 
Dim C_ReqTypeCd   '���ſ�û���� 
Dim C_ReqTypeNm   '���ſ�û���и� 
Dim C_ORGCd       '�������� 
Dim C_ORGNm       '���������� 
Dim C_MrpRunNo    'MRP

Dim C_SpplCd                 '����ó 
Dim C_SpplPopup              '����ó �˾� 
Dim C_SpplNm 	             '����ó�� 
Dim C_Quota_Rate             '��к��� 
Dim C_ApportionQty           '��η� 
Dim C_PlanDt                 '���ֿ����� 
Dim C_GrpCd 	             '���ű׷� 
Dim C_GrpPopup               '���ű׷��˾� 
Dim C_GrpNm 	             '���ű׷�� 
Dim C_ORGCd2
Dim C_ParentPrNo 	         '���� ��û��ȣ (Ű��)
Dim C_ParentRowNo            '���� row ��ȣ 
Dim C_Flag                   '�ڱ� ��ȣ 

'/* 9�� ������ġ: ����� ���� Į����ġ ���� 7 --> 6  - START */
Dim lgIntFlgModeM           'Variable is for Operation Status
Dim lglngHiddenRows()		'Multi���� �������� ���� ����	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
Dim lgSortKey1
Dim lgSortKey2
Dim IsOpenPop
Dim lgCurrRow
Dim lgSpdHdrClicked	'2003-03-01 Release �߰� 
Dim lgPageNo1
Dim EndDate, StartDate,CurrDate, iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"
CurrDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate = UnIDateAdd("m", 1, CurrDate, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, CurrDate, parent.gDateFormat)

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgStrPrevKey=""
    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
    lgSortKey2 = 2
    lgPageNo1 = 0
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.Value = Parent.gPlant
' 	frm1.txtFrDt.Text=StartDate
' 	frm1.txtToDt.Text=EndDate

	Call Setminorcd()
	Call SetToolbar("1100000000001111")
	frm1.btnAutoSel.disabled = True
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
    frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
	Set gActiveSpdSheet = frm1.vspdData
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1

		.vspdData.ReDraw = false

		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20040425",, parent.gAllowDragDropSpread

	   .vspdData.MaxCols = C_MrpRunNo+1
	   .vspdData.MaxRows = 0

		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetCheck    C_CfmFlg		, "����"		, 8,,,true
		ggoSpread.SSSetEdit		C_ReqNo			, "��û��ȣ"	, 18
		ggoSpread.SSSetEdit		C_PlantCd		, "����"		, 10
		ggoSpread.SSSetEdit		C_PlantNm		, "�����"		, 20
		ggoSpread.SSSetEdit		C_ItemCd		, "ǰ��"		, 18
		ggoSpread.SSSetEdit		C_ItemNm		, "ǰ���"		, 25
		ggoSpread.SSSetEdit 	C_SpplSpec		, "ǰ��԰�"	, 20
		SetSpreadFloatLocal		C_ReqQty		, "��û��"		, 15, 1,3
		ggoSpread.SSSetEdit		C_Unit			, "����"		, 6
		ggoSpread.SSSetEdit		C_TrackingNo	, "Tracking No"	, 18
		ggoSpread.SSSetDate		C_ReqDt			, "�ʿ���"		, 10, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit		C_ReqStateCd	, "���ſ�û����",12
		ggoSpread.SSSetEdit		C_ReqStateNm	, "���ſ�û���¸�",14
		ggoSpread.SSSetEdit		C_ReqTypeCd		, "���ſ�û����",12
		ggoSpread.SSSetEdit		C_ReqTypeNm		, "���ſ�û���и�",14
		ggoSpread.SSSetEdit		C_ORGCd			, "��������"		, 10,,,4,2
        ggoSpread.SSSetEdit		C_ORGNm			, "����������"		, 20
		ggoSpread.SSSetEdit		C_MrpRunNo		, "MRP Run��ȣ"	,20

		Call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantNm)
	    Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)
		Call ggoSpread.MakePairsColumn(C_ORGCd,C_ORGNm)
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols,	.vspdData.MaxCols,	True)
		.vspdData.ReDraw = True
    End With

	Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet2
' Function Desc : This method initializes spread sheet2 column property
'=======================================================================================================
Sub InitSpreadSheet2()
	Call InitSpreadPosVariables2()
    With frm1
		.vspdData2.ReDraw = false

		ggoSpread.Source = frm1.vspdData2
        ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread

	   .vspdData2.MaxCols = C_Flag + 1
	   .vspdData2.MaxRows = 0

		Call GetSpreadColumnPos("B")

		ggoSpread.SSSetEdit		C_SpplCd		, "����ó"			, 10,,,10,2
		ggoSpread.SSSetButton	C_SpplPopup
		ggoSpread.SSSetEdit 	C_SpplNm		, "����ó��"		, 18
		SetSpreadFloatLocal		C_Quota_Rate	, "��к���(%)"		, 12,1,5
		SetSpreadFloatLocal		C_ApportionQty	, "��η�"			, 12,1,3
		ggoSpread.SSSetDate		C_PlanDt		, "���ֿ�����"		, 12, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit 	C_GrpCd			, "���ű׷�"		, 10,,,4,2
		ggoSpread.SSSetButton	C_GrpPopUp
		ggoSpread.SSSetEdit 	C_GrpNm			, "���ű׷��"		, 20
		ggoSpread.SSSetEdit		C_ORGCd2		, "��������"		, 10,,,4,2
        ggoSpread.SSSetEdit 	C_ParentPrNo	, "��û��ȣ"		, 20
		ggoSpread.SSSetEdit     C_ParentRowNo	, ""			, 5,2,,,2
		ggoSpread.SSSetEdit     C_Flag			, ""			, 5,2,,,2

		Call ggoSpread.MakePairsColumn(C_SpplCd,C_SpplNm)
		Call ggoSpread.MakePairsColumn(C_GrpCd,C_GrpNm)

		Call ggoSpread.SSSetColHidden(C_ORGCd2,C_Flag,True)
		Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols,.vspdData2.MaxCols ,	True)

		.vspdData2.ReDraw = True
    End With

	Call SetSpreadLock2()
	Call SetSpreadColor2(-1,-1)
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False

    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLock 1 , -1
    ggoSpread.spreadUnlock C_CfmFlg, -1,C_CfmFlg, -1
    .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadLock2()
    With frm1
    .vspdData2.ReDraw = False
    ggoSpread.Source = frm1.vspdData2

	ggoSpread.SpreadLock		C_SpplCd,			-1,	C_SpplNm,		-1
	ggoSpread.spreadUnlock		C_Quota_Rate,		-1,	C_PlanDt,	-1
	ggoSpread.SSSetRequired		C_Quota_Rate,		-1, -1
	ggoSpread.SSSetRequired		C_ApportionQty,		-1, -1
	ggoSpread.SSSetRequired		C_PlanDt,			-1, -1
	ggoSpread.spreadUnlock		C_GrpCd,			-1,	C_GrpPopup,    -1
	ggoSpread.SSSetRequired		C_GrpCd,			-1, -1
	ggoSpread.SpreadLock		C_GrpNm,			-1,	C_GrpNm,		-1

	.vspdData2.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False

	ggoSpread.SSSetProtected C_ReqNo,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PlantCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_PlantNm,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemCd,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemNm,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_SpplSpec,	    pvStartRow, pvEndRow

	ggoSpread.SSSetProtected C_ReqQty,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Unit,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_TrackingNo,		pvStartRow, pvEndRow

	ggoSpread.SSSetProtected C_ReqDt,			pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ReqStateCd,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ReqStateNm,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ReqTypeCd,		pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ReqTypeNm,		pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_ORGCd,			pvStartRow,	pvEndRow
    ggoSpread.SSSetProtected C_ORGNm,		    pvStartRow,	pvEndRow
	ggoSpread.SSSetProtected C_MrpRunNo,		pvStartRow, pvEndRow

    .vspdData.ReDraw = True

    End With
End Sub


Sub SetSpreadColor2(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData2.ReDraw = False

	ggoSpread.SSSetProtected C_SpplCd,		    pvStartRow,	pvEndRow
	ggoSpread.SSSetProtected C_SpplPopup,		pvStartRow,	pvEndRow
	ggoSpread.SSSetProtected C_SpplNm,		    pvStartRow,	pvEndRow
	ggoSpread.SSSetRequired  C_Quota_Rate,		pvStartRow,	pvEndRow
    ggoSpread.SSSetRequired  C_ApportionQty,	pvStartRow,	pvEndRow
    ggoSpread.SSSetRequired  C_PlanDt,			pvStartRow,	pvEndRow
    ggoSpread.SSSetRequired  C_GrpCd,			pvStartRow,	pvEndRow
    ggoSpread.SSSetProtected C_GrpNm,		    pvStartRow,	pvEndRow
    ggoSpread.SSSetProtected .vspdData2.MaxCols, pvStartRow, pvEndRow
   .vspdData2.ReDraw = True

    End With
End Sub

Sub SetSpreadColor3(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    .vspdData2.ReDraw = False

	ggoSpread.spreadUnlock	 C_SpplCd,		pvStartRow,	C_SpplCd,	pvStartRow
	ggoSpread.SSSetRequired  C_SpplCd,		pvStartRow,	pvStartRow
	ggoSpread.spreadUnlock	 C_SpplPopup,		pvStartRow,	C_SpplCd,	pvStartRow
	ggoSpread.SSSetProtected C_SpplNm,		    pvStartRow,	pvEndRow
	ggoSpread.SSSetRequired  C_Quota_Rate,		pvStartRow,	pvEndRow
    ggoSpread.SSSetRequired  C_ApportionQty,	pvStartRow,	pvEndRow
    ggoSpread.SSSetRequired  C_PlanDt,			pvStartRow,	pvEndRow
    ggoSpread.SSSetRequired  C_GrpCd,			pvStartRow,	pvEndRow
    ggoSpread.SSSetProtected C_GrpNm,		    pvStartRow,	pvEndRow

   .vspdData2.ReDraw = True

    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_CfmFlg        = 1			 '���� 
	C_ReqNo         = 2			 '��û��ȣ 
	C_PlantCd       = 3			 '���� 
	C_PlantNm       = 4			 '����� 
	C_ItemCd        = 5			 'ǰ�� 
	C_ItemNm        = 6			 'ǰ��� 
	C_SpplSpec      = 7			 '�԰� 
	C_ReqQty        = 8			 '��û�� 
	C_Unit          = 9			 '���� 
	C_TrackingNo	= 10
	C_ReqDt         = 11		 '�ʿ��� 
	C_ReqStateCd    = 12		 '���ſ�û���� 
	C_ReqStateNm    = 13		 '���ſ�û���¸� 
	C_ReqTypeCd     = 14		 '���ſ�û���� 
	C_ReqTypeNm     = 15		 '���ſ�û���и� 
	C_ORGCd         = 16          '�������� 
	C_ORGNm         = 17          '���������� 
	C_MrpRunNo      = 18		 'MRP
End Sub

'==========================================  2.2.7 InitSpreadPosVariables2()  =============================
' Function Name : InitSpreadPosVariables2
' Function Desc : This method Assigns Sequential Number to spread sheet2 column
'=========================================================================================================
Sub InitSpreadPosVariables2()
	C_SpplCd        = 1          '����ó 
	C_SpplPopup     = 2          '����ó �˾� 
	C_SpplNm 	    = 3          '����ó�� 
	C_Quota_Rate    = 4          '��к��� 
	C_ApportionQty  = 5          '��η� 
	C_PlanDt        = 6          '���ֿ����� 
	C_GrpCd 	    = 7         '���ű׷� 
	C_GrpPopup      = 8         '���ű׷��˾� 
	C_GrpNm 	    = 9         '���ű׷�� 
	C_ORGCd2		= 10
	C_ParentPrNo    = 11	     '���� ��û��ȣ (Ű��)
	C_ParentRowNo   = 12         '���� row ��ȣ 
	C_Flag          = 13         '�ڱ� ��ȣ 
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
				C_CfmFlg		=	iCurColumnPos(1)      '���� 
				C_ReqNo		    =	iCurColumnPos(2)      '��û��ȣ 
				C_PlantCd       =	iCurColumnPos(3)      '���� 
				C_PlantNm       =	iCurColumnPos(4)      '����� 
				C_ItemCd        =	iCurColumnPos(5)      'ǰ�� 
				C_ItemNm        =	iCurColumnPos(6)      'ǰ��� 
				C_SpplSpec      =	iCurColumnPos(7)      '�԰� 
				C_ReqQty        =	iCurColumnPos(8)      '��û�� 
				C_Unit          =	iCurColumnPos(9)      '���� 
				C_TrackingNo	=	iCurColumnPos(10)
				C_ReqDt         =	iCurColumnPos(11)     '�ʿ��� 
				C_ReqStateCd    =	iCurColumnPos(12)     '���ſ�û���� 
				C_ReqStateNm    =	iCurColumnPos(13)     '���ſ�û���¸� 
				C_ReqTypeCd     =	iCurColumnPos(14)     '���ſ�û���� 
				C_ReqTypeNm     =	iCurColumnPos(15)     '���ſ�û���и� 
				C_ORGCd         =	iCurColumnPos(16)     '�������� 
				C_ORGNm         =	iCurColumnPos(17)     '���������� 
				C_MrpRunNo      =	iCurColumnPos(18)     'MRP

	   Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

				C_SpplCd        =	iCurColumnPos(1)         '����ó 
				C_SpplPopup     =	iCurColumnPos(2)         '����ó �˾� 
				C_SpplNm 	    =	iCurColumnPos(3)         '����ó�� 
				C_Quota_Rate    =	iCurColumnPos(4)         '��к��� 
				C_ApportionQty  =	iCurColumnPos(5)         '��η� 
				C_PlanDt        =	iCurColumnPos(6)         '���ֿ����� 
				C_GrpCd 	    =	iCurColumnPos(7)        '���ű׷� 
				C_GrpPopup      =	iCurColumnPos(8)        '���ű׷��˾� 
				C_GrpNm 	    =	iCurColumnPos(9)        '���ű׷�� 
				C_ORGCd2		=   iCurColumnPos(10)        '���ű׷�� 
				C_ParentPrNo    =	iCurColumnPos(11)	     '���� ���� (Ű��)
				C_ParentRowNo   =	iCurColumnPos(12)        '���� row ��ȣ 
				C_Flag          =	iCurColumnPos(13)        '�ڱ� ��ȣ 
	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : SpreadItem PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "����"
	arrParam(1) = "B_Plant"

	arrParam(2) = Trim(frm1.txtPlantCd.Value)

	arrParam(4) = ""
	arrParam(5) = "����"

	arrField(0) = "Plant_Cd"
	arrField(1) = "Plant_NM"

	arrHeader(0) = "����"
	arrHeader(1) = "�����"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)
		frm1.txtPlantNm.Value= arrRet(1)
		frm1.txtPlantCd.focus
	End If
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenItem()
' Description : SpreadItem PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
    Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If  Trim(frm1.txtPlantCd.Value) = "" Then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

    iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtitemcd.focus
		Exit Function
	Else
		frm1.txtitemcd.Value    = arrRet(0)
		frm1.txtitemNm.Value    = arrRet(1)
		frm1.txtitemcd.focus
	End If
End Function

'===========================================================================
' Function Name : OpenMrp
' Function Desc : OpenMrp Reference Popup
'===========================================================================
Function OpenMrp()
    Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

    If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	IsOpenPop = True

	arrParam(0) = "MRP Run��ȣ"				<%' �˾� ��Ī %>
	arrParam(1) = "(select distinct a.order_no A,a.confirm_dt B," & FilterVar("������������", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_EXPL_HISTORY a, m_pur_req b where a.order_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "
    arrParam(1) = arrParam(1) & "union "
    arrParam(1) = arrParam(1) & "select distinct  a.run_no A, a.start_dt B ," & FilterVar("MRP����", "''", "S") & " D "
    arrParam(1) = arrParam(1) & "from P_MRP_HISTORY a, m_pur_req b where a.run_no = b.mrp_run_no and a.plant_cd = b.plant_cd and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & ") as g" <%' TABLE ��Ī %>


	arrParam(2) = Trim(frm1.txtMRP.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "MRP Run��ȣ"				<%' TextBox ��Ī %>

	arrField(0) = "A"
	arrField(1) = "B"
	arrField(2) = "D"

	arrHeader(0) = "MRP Run��ȣ"				<%' Header��(0)%>
	arrHeader(1) = "����"					<%' Header��(1)%>
	arrHeader(2) = "��������"				<%' Header��(2)%>

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtMRP.focus
		Exit Function
	Else
		frm1.txtMRP.value = arrRet(0)
		frm1.txtMRP.focus
	End If

End Function

'===========================================================================
' Function Name : OpenSoNo
' Function Desc : OpenSoNo Reference Popup
'===========================================================================
Function OpenSoNo()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	iCalledAspName = AskPRAspName("S3111PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,""), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		frm1.txtSoNo.focus
		Exit Function
	Else
		frm1.txtSoNo.value = strRet
		frm1.txtSoNo.focus
	End If
End Function


'------------------------------------------  OpenTrackingNo()  -------------------------------------------------
'	Name : OpenTrackingNo()
'	Description : TrackNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = Trim(frm1.txtPlantCd.value)	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 

	iCalledAspName = AskPRAspName("S3135PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = Trim(arrRet)
		frm1.txtTrackNo.focus
	End If
End Function

'------------------------------------------  OpenSSupplier()  -------------------------------------------------
' Name : OpenSSupplier()
' Description : SpreadItem PopUp
'--------------------------------------------------------------------------------------------------------- ----
Function OpenSSupplier()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

	arrParam(0) = "����ó"
	arrParam(1) = "B_Biz_Partner"

	frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col = C_SpplCd
	arrParam(2) = FilterVar(Trim(frm1.vspdData2.text)," ","SNM")

	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "����ó"

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = "����ó"
	arrHeader(1) = "����ó��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData2
			.Row = .ActiveRow
			.Col = C_SpplCd
			.text = arrRet(0)
			.Row = .ActiveRow
			.Col = C_SpplNm
			.text = arrRet(1)
			Call SpplChange()
		End With
	End If

End Function

'------------------------------------------  OpenSGrp()  -------------------------------------------------
' Name : OpenSGrp()
' Description : SpreadItem PopUp
'--------------------------------------------------------------------------------------------------------- ----
Function OpenSGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.vspdData2.Col=C_GrpCd
	frm1.vspdData2.Row=frm1.vspdData2.ActiveRow

	arrParam(0) = "���ű׷�"
	arrParam(1) = "B_pur_grp"
	arrParam(2) = Trim(frm1.vspdData2.Text)
	arrParam(3) = ""

	frm1.vspdData2.Row=frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col=C_ORGCd2

	arrParam(4) = "Usage_flg=" & FilterVar("Y", "''", "S") & "  and PUR_ORG =  " & FilterVar(frm1.vspdData2.Text, "''", "S") & " "
	arrParam(5) = "���ű׷�"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "���ű׷�"
    arrHeader(1) = "���ű׷��"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData2
			.Row = .ActiveRow
			.Col = C_GrpCd
			.text = arrRet(0)
			.Row = .ActiveRow
			.Col = C_GrpNm
			.text = arrRet(1)
			Call vspdData2_Change(C_GrpCd,.ActiveRow)
		End With
	End If
End Function

'------------------------------------------  OpenORG()  -------------------------------------------------
'	Name : OpenORG()
'	Description : OpenORG PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSORG()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"
	arrParam(1) = "B_Pur_Org"

	frm1.vspdData2.Row=frm1.vspdData2.ActiveRow
	frm1.vspdData2.Col=C_ORGCd

	arrParam(2) = Trim(frm1.vspdData2.Text)
'	arrParam(3) = Trim(frm1.txtORGNm.Value)

	arrParam(4) = ""
	arrParam(5) = "��������"

    arrField(0) = "PUR_ORG"
    arrField(1) = "PUR_ORG_NM"

    arrHeader(0) = "��������"
    arrHeader(1) = "����������"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData2
			.Row = .ActiveRow
			.Col = C_ORGCd
			.text = arrRet(0)
			.Row = .ActiveRow
			.Col = C_ORGNm
			.text = arrRet(1)
			.Row = .ActiveRow
			.Col = C_GrpCd
			.text = ""
			.Row = .ActiveRow
			.Col = C_GrpNm
			.text = ""
		End With
	End If
End Function

'------------------------------------------  SetSpreadFloatLocal()  --------------------------------------
' Name : SetSpreadFloatLocal()
' Description :
'---------------------------------------------------------------------------------------------------------
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )

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
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'=======================================================================================================
Function DbQuery2(ByVal Row, Byval NextQueryFlag)
	DbQuery2 = False

	Dim strVal
	Dim lngRet
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iStrOrgCd
	Dim pRow

	'/* 9�� ������ġ: ���� ���������� �ణ �̵� �� �̹� ��ȸ�� �ڷᳪ �Էµ� �ڷḦ �о� ���� ������ '' â ���� - START */
	Call LayerShowHide(1)

	With frm1

		.vspdData.Row = Row
		.vspdData.Col = C_OrgCd
		iStrOrgCd = .vspdData.text

		.vspdData.Col = .vspdData.MaxCols
		pRow = CInt(.vspdData.Text)

		If lglngHiddenRows(pRow - 1) <> 0 And NextQueryFlag = False Then

			.vspdData2.ReDraw = False

			 lngRet = ShowFromData(pRow, lglngHiddenRows(pRow - 1))	'ex) ù��° �׸����� Ư�� Row�� �ش��ϴ� �ι�° �׸����� Row���� 10���϶� ������ �����Ͱ� 3��° ���� 6��°���� 4���̸� 3�� �����ϴ� ����� �����ϴ� �Լ���.

			Call SetToolBar("11001111001011")
			Call LayerShowHide(0)

			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()

			.vspdData2.ReDraw = True
			DbQuery2 = True
			Exit Function
		End If


		.vspdData.Row = CInt(Row)
		.vspdData.Col = C_ReqNo

		strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		strVal = strVal & "&txtPrNo=" & Trim(.vspdData.text)
		strVal = strVal & "&txtOgrCd=" & Trim(iStrOrgCd)
		strVal = strVal & "&lgPageNo1="		 & lgPageNo1						'��: Next key tag
		strVal = strVal & "&lglngHiddenRows=" & lglngHiddenRows(Row - 1)
		strVal = strVal & "&lRow=" & CStr(pRow)

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
	Dim totalquotarate

	With frm1.vspdData2

		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = C_Flag

		.Col2 = C_Flag
		.DestCol = 0
		.DestRow = lngRangeFrom
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False

		'��ȸ�ϰ� ���� �ش� 2nd Grid�� ��Ŀ�� ����������. ��Ŀ�� ��ġ���� Insert Row�� �Ա⶧�� 200308
		.Col = C_SpplCd
		.Row = lngRangeTo
		.Action = 0

	End With

	For index = 1 to frm1.vspdData2.MaxRows
		Call checkdt(index)
	Next

	frm1.vspdData.Focus

	DbQueryOk2 = true
End Function

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
'	Description :
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
        end if
    End With
End Sub


'------------------------------------  Setretflg()  ----------------------------------------------
'	Name : Setreference()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub Setminorcd()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iminorcd

    Err.Clear

	Call CommonQueryRs(" minor_cd ", " b_configuration ", " major_cd = " & FilterVar("M2105", "''", "S") & " and reference = " & FilterVar("Y", "''", "S") & "  and seq_no = " & FilterVar("1", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iminorcd = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation, parent.gLogoName
		Err.Clear
		Exit Sub
	End If

    If Trim(lgF0) <> "" Then
        If UCase(Trim(iminorcd(0))) = "D" Then
            frm1.rdoAssflg(0).Checked = True
        ElseIf UCase(Trim(iminorcd(0))) = "R" Then
            frm1.rdoAssflg(1).Checked = True
        Else
            frm1.rdoAssflg(2).Checked = True
        End If
    End If
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

	With frm1.vspdData2
		.Row		= Row
		.Col		= C_ParentRowNo
		iparentrow  = Trim(.text)

		.Col		= C_Quota_Rate
		iquotarate  = Unicdbl(.text)

		lngRangeFrom = DataFirstRow(iparentrow)
	    lngRangeTo   = DataLastRow(iparentrow)

		totalquotarate = 0
		totalApportionQty = 0

		For index = lngRangeFrom  to lngRangeTo
		    .Row = index
		    .Col = 0
		    if Trim(.Text) <> ggoSpread.DeleteFlag  then
				.Col = C_Quota_Rate
				totalquotarate = totalquotarate + Unicdbl(.text)
		        if index <> clng(Row) then
				    .Col = C_ApportionQty
				    totalApportionQty = totalApportionQty + Unicdbl(.text)
		        end if
		    end if
		Next

		frm1.vspdData.Row = iparentrow
		frm1.vspdData.Col = C_ReqQty
		iReqQty = Unicdbl(frm1.vspdData.text)

		'�հ� ������� 100�̸� ��η� = ��û�� - �����η��� 
		if totalquotarate = 100 then
		    iApportionQty = iReqQty - totalApportionQty
		else
			iApportionQty = (iquotarate * iReqQty)/100
	    end if

		.Row  = Row
		.Col  = C_ApportionQty
		.text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
		.Col = C_Flag
		.text = ggoSpread.UpdateFlag
	End with

End Sub

'==========================================   QuotaRateChange()  ======================================
'	Name : QuotaRateChange()
'	Description :
'=================================================================================================
Sub QuotaRateChange(ByVal Row)
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

		.Col		= C_ApportionQty
		iApportionQty  = Unicdbl(.text)

		lngRangeFrom = DataFirstRow(iparentrow)
	    lngRangeTo   = DataLastRow(iparentrow)

		totalquotarate = 0
		totalApportionQty = 0

		for index = lngRangeFrom  to lngRangeTo
		    .Row = index
		    .Col = 0
		    if Trim(.Text) <> ggoSpread.DeleteFlag  then

				if index <> clng(Row) then
					.Col = C_Quota_Rate
					totalquotarate = totalquotarate + Unicdbl(.text)
		        end if

		    end if
		next

		frm1.vspdData.Row = iparentrow
		frm1.vspdData.Col = C_ReqQty
		iReqQty = Unicdbl(frm1.vspdData.text)

		'�հ� ������� 100�̸� ��η� = ��û�� - �����η��� 
		if totalApportionQty = iReqQty then
		    totalquotarate = 100
		else
			totalquotarate = (iApportionQty * 100) / iReqQty
	    end if

		.Row  = Row
		.Col  = C_Quota_Rate
		.text = UNIFormatNumber(totalquotarate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

		.Col = C_Flag
		.text = ggoSpread.UpdateFlag
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
    Dim strssText1, strssText2, strssText3
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
		End if

	End with

	lngRangeFrom = DataFirstRow(iparentrow)
	lngRangeTo   = DataLastRow(iparentrow)

	for index = lngRangeFrom to lngRangeTo
	    if index <> iRow and strssText2 <> "" then
	        frm1.vspdData2.Row = index
	        frm1.vspdData2.Col = C_SpplCd
	        if UCase(strssText2) = UCase(Trim(frm1.vspdData2.text)) then
                Call DisplayMsgBox("17A005", "X","{" & strssText1 & "}", "{" & strssText2 & "}")
				Call LayerShowHide(0)
				frm1.vspdData2.Row = iRow
	            frm1.vspdData2.Col = C_SpplCd
	            frm1.vspdData2.text = ""
 	            Exit sub
	        end if
	    end if
	next

    strVal = BIZ_PGM_ID & "?txtMode=" & "LookSppl"
    strVal = strVal & "&txtPrNo=" & strssText1
    strVal = strVal & "&txtBpCd=" & strssText2

    If LayerShowHide(1) = False Then Exit Sub

	Call RunMyBizASP(MyBizASP, strVal)

End Sub

'==========================================   GroupChange()  ======================================
'	Name : GroupChange()
'	Description :
'=================================================================================================
Sub GroupChange()
    Err.Clear
End Sub

'=======================================================================================================
'   Function Name : ShowFromData
'   Function Desc :
'=======================================================================================================
Function ShowFromData(Byval Row, Byval lngShowingRows)
	ShowFromData = 0

	Dim lngRow
	Dim lngStartRow

	With frm1.vspdData2

		Call SortSheet()
		'------------------------------------
		' Find First Row
		'------------------------------------
		lngStartRow = 0

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
			'ī�Ǿ׼� ���� ����. - ����׸��� ���������̵��� �ϴܱ׸��� Update Flag�� �̻������� ���󶧹�.200308
			'.Action = 19	'SS_ACTION_COPY_RANGE
			.RowHidden = False

			.BlockMode = False

			If lngStartRow > 1 Then
				.BlockMode = True
				.Row = 1
				.Row2 = lngStartRow - 1
				.RowHidden = True
				.BlockMode = False
			End If

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

			'Sppl Cd�� ��Ŀ�� ������ ���� 200308
			.Row = lngStartRow	'2003-03-01 Release �߰� 
			.Col = C_SpplCd		'2003-03-01 Release �߰� 
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

	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			.Col = 0
			If .Text = strInsertMark Or .Text = strDeleteMark Or .Text = strUpdateMark Then
				ChangeCheck = True
			End If
		Next
	End With
End Function

'=======================================================================================================
' Function Name : CheckDataExist
' Function Desc :
'=======================================================================================================
Function CheckDataExist()
	CheckDataExist = False
	Dim i

	If frm1.vspdData2.MaxRows = 0 Then
		With frm1.vspdData
			.Row = .ActiveRow
		    .Col = C_CfmFlg
			If .value = 1 AND frm1.vspdData2.RowHidden = False Then
				CheckDataExist = True
				Exit Function
			End If
		End With
	Else
		With frm1.vspdData2
			For i = 1 To .MaxRows
				.Row = i
				If .RowHidden = False Then
					CheckDataExist = True
					Exit Function
				End IF
			Next
		End With
	End IF
End Function

'=======================================================================================================
' Function Name : ShowDataFirstRow
' Function Desc :
'=======================================================================================================
Function ShowDataFirstRow()
	ShowDataFirstRow = 0
	Dim i

	With frm1.vspdData
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
' Function Name : ShowDataFirstRow2
' Function Desc :
'=======================================================================================================
Function ShowDataFirstRow2()
	ShowDataFirstRow2 = 0
	Dim i

	With frm1.vspdData2
		For i = 1 To .MaxRows
			.Row = i
			If .RowHidden = False Then
				ShowDataFirstRow2 = i
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

	With frm1.vspdData
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
' Function Name : ShowDataLastRow2
' Function Desc :
'=======================================================================================================
Function ShowDataLastRow2()
	ShowDataLastRow2 = 0
	Dim i

	With frm1.vspdData2
		For i = .MaxRows To 1 Step -1
			.Row = i
			If .RowHidden = False Then
				ShowDataLastRow2 = i
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
			'.Col = C_Flag
			'.Text = strMark
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

 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    Else
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
    End IF

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
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
 		lgSpdHdrClicked = 0		'2003-03-01 Release �߰� 
 		Call Sub_vspdData_ScriptLeaveCell(0, 0, Col, frm1.vspdData.ActiveRow, False)
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
 	Dim lngStartRow
 	Dim i,k
 	Dim strFlag,strFlag1
 	Dim iActiveRow

 	gMouseClickStatus = "SP2C"

 	Set gActiveSpdSheet = frm1.vspdData2

 	If Row <= 0 Then
 		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    Else
		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
    End IF

 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 AND Col <> 0 Then	'2003-03-01 Release �߰� 
 		ggoSpread.Source = frm1.vspdData2

 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
 		frm1.vspdData.Col = frm1.vspdData.MaxCols

 		iActiveRow = frm1.vspdData.Text

 		frm1.vspdData2.Redraw = False

		lngStartRow = CInt(ShowFromData(iActiveRow, CInt(lglngHiddenRows(iActiveRow - 1))))
		frm1.vspdData2.Redraw = True
		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Ascending
 			lgSortKey2 = 2
 		ElseIf lgSortKey2 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey2, lngStartRow, lngStartRow + CInt(lglngHiddenRows(iActiveRow - 1)) - 1	'Sort in Descending
 			lgSortKey2 = 1
		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If

 	With frm1.vspdData2
 		For i = 1 to .MaxRows
 			.Row = i
 			.col = 0
 			If .Rowhidden = False Then
 				k = K + 1
 				if .text <> ggoSpread.InsertFlag  AND .text <> ggoSpread.UpdateFlag AND .text <> ggoSpread.deleteFlag then
 					.text = k
 				end if
 			End If
 		Next
 	End With

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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData2_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()	'###�׸��� ������ ���Ǻκ�###
	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
'	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
	Call InitSpreadSheet                                                        'Setup the Spread sheet1
	Call InitSpreadSheet2
	Call InitVariables
	Call SetDefaultVal
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
		Next

		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet2()
		frm1.vspdData2.Redraw = False

		Call ggoSpread.ReOrderingSpreadData("F")

		Call DbQuery2(frm1.vspdData.ActiveRow,False)

		lngRangeFrom = Clng(ShowDataFirstRow)
		lngRangeTo = Clng(ShowDataLastRow)

		lRow = frm1.vspdData.ActiveRow	'###�׸��� ������ ���Ǻκ�###
		frm1.vspdData2.Redraw = True
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo lRow
    End If

 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release �߰� 
		Exit Sub
	End If
	if frm1.vspddata.row = 0 then exit sub
	Call Sub_vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)
End Sub

'=======================================================================================================
'   Event Name : Sub_vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub Sub_vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	Dim lRow
	if Row = 0 then exit sub
	If Row <> NewRow And NewRow > 0 Then
		If CheckRunningBizProcess = True Then
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = 1
			frm1.vspdData.Action = 0
			Exit Sub
		End If
		'/* �ٸ� �۾��� �̷������ ��Ȳ���� �ٸ� �� �̵� �� ��ȸ�� �̷�� ���� �ʵ��� �Ѵ�. - END */
		lgCurrRow = NewRow
'		lgIntFlgModeM = Parent.OPMD_CMODE

		frm1.vspdData2.ReDraw = False
		frm1.vspdData2.BlockMode = True
		frm1.vspdData2.Row = 1
		frm1.vspdData2.Row2 = frm1.vspdData2.MaxRows
		frm1.vspdData2.RowHidden = True
		frm1.vspdData2.BlockMode = False
		frm1.vspdData2.ReDraw = True

		If DbQuery2(lgCurrRow, False) = False Then	Exit Sub
	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
	Dim strMark
	Dim iparentrow
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

	With frm1.vspdData2
		.Row = Row
		.Col = C_ParentRowNo
		iparentrow = .text
		.Col = 0
		strMark = .Text

		.Col = C_Flag
		.Text = strMark

		Select Case Col
	        Case C_PlanDt
				 .Row = Row
				 .Col = Col

				 If UniConvDateToYYYYMMDD(frm1.vspdData2.Text,parent.gDateFormat,"") < UniConvDateToYYYYMMDD(CurrDate,parent.gDateFormat,"") Then
				     Call sprRedComColor(C_PlanDt,Row,Row)
				 else
				     Call sprBlackComColor(C_PlanDt,Row,Row)
				 end if
	        Case C_SpplCd
	             Call SpplChange()

	        Case C_Quota_Rate
	             Call ApportionQtyChange(Row)
	        Case C_GrpCd
				 Call GroupChange()
			Case C_ApportionQty
				 Call QuotaRateChange(Row)
	    end select

    End With

    With frm1.vspdData
		' === 2005.07.13 Tracking No. 9956 ====================
		.Row = .ActiveRow
'		.Row = iparentrow
		' === 2005.07.13 Tracking No. 9956 ====================

		.Col = C_CfmFlg

		If .value = 0 then
		    .value = 1
		end if

		.Col = 0
		.Text = ggoSpread.UpdateFlag
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

    '/* �ػ󵵿� ������� �������ǵ��� ���� - START */
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	        '��: ������ üũ 
    '/* �ػ󵵿� ������� �������ǵ��� ���� - END */
		If lgStrPrevKey <> ""   Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If

			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
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
    Dim lConvRow

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    With frm1

    	lRow = .vspdData.ActiveRow
    	'/* �ػ󵵿� ������� �������ǵ��� ���� - START */
    	If ShowDataLastRow < NewTop + VisibleRowCnt(.vspdData2, NewTop) Then	        '��: ������ üũ 
		'/* �ػ󵵿� ������� �������ǵ��� ���� - END */
    		If lgPageNo1 > 0 Then            '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_CfmFlg And Row > 0 Then
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
'======================================================================================================
'   Event Name : vspdData2_ButtonClicked
'   Event Desc :
'=======================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData2

		ggoSpread.Source = frm1.vspdData2

		If Row > 0 And Col = C_SpplPopup Then
			Call OpenSSupplier()
		Elseif Row > 0 And Col = C_GrpPopup Then
			Call OpenSGrp()
		End if

	End With
End Sub

'==========================================================================================
'   Event Name : txtFrDt
'   Event Desc :
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtFrDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtToDt
'   Event Desc :
'==========================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtToDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'==========================================================================================
Sub txtPoFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtPoFrDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtToDt
'   Event Desc :
'==========================================================================================
Sub txtPoToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtPoToDt.Focus
	End if
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
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPoToDt_KeyDown(KeyCode, Shift)
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
'    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
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
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then
			Call DisplayMsgBox("17a003", "X","�ʿ���", "X")
			Exit Function
		End if

		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) and Trim(.txtPoFrDt.text)<>"" and Trim(.txtPoToDt.text)<>"" then
			Call DisplayMsgBox("17a003", "X","���ֿ�����", "X")
			Exit Function
		End if

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
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew()
    FncNew = False

    Dim IntRetCD

	'-----------------------
    'Check previous data area
    '-----------------------
    If ChangeCheck = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

	'-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
'    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables                                                      'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.ClearSpreadData

	ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData

    Call SetDefaultVal

	Set gActiveElement = document.activeElement
    FncNew = True
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncDelete()
    Dim lDelRows
    Dim iDelRowCnt, i
	if frm1.vspdData.Maxrows < 1 then exit function

    With frm1.vspdData
    	.focus
		ggoSpread.Source = frm1.vspdData
        lDelRows = ggoSpread.DeleteRow

    End With
	Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave()
    FncSave = False

    Dim IntRetCD


    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

    ggoSpread.Source = frm1.vspdData

    '-----------------------
    'Precheck area
    '-----------------------
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

    'ȭ�鿡 ���̴� ���� �������忡 ���߰� �Ǿ����� Hidden �������忡 �ݿ��� �ȵ� �� üũ START
    If DefaultCheck = False Then
    	Exit Function
    End If

    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then
       		Exit Function
    End If
    If Not chkField(Document, "2") Then
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

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
	FncCopy = false

	Dim IntRetCD
	Dim lRow
	Dim lRow2
	Dim totalQuotaRate,totalApportionQty
	Dim iQuotaRate,iApportionQty,iReqQty

	With frm1
		'Check Spread2 Data Exists for the keys

		If CheckDataExist = False Then
			Exit function
		End If

    	.vspdData2.ReDraw = False

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.CopyRow

		lRow2 = .vspdData2.ActiveRow
		.vspdData2.Row = lRow2

		.vspdData2.Col = C_SpplCd
		.vspdData2.Text = ""

		.vspdData2.Col = C_SpplNm
		.vspdData2.Text = ""

		.vspdData2.Col = C_Quota_Rate
		.vspdData2.Text = 0

	    .vspdData2.Col = C_ApportionQty
		.vspdData2.Text = 0

		.vspdData2.Col = C_Flag
		.vspdData2.Text = ggoSpread.InsertFlag

		Call SetSpreadColor3(lRow2, lRow2)

	    lRow = .vspdData.ActiveRow
	    .vspdData.Row = lRow
	    .vspdData.Col = C_ReqQty
        iReqQty = Unicdbl(.vspdData.text)

		'�������� ���� �ش� Ű�� ���� Client�� Data Row���� ������ 
		lglngHiddenRows(lRow - 1) = lglngHiddenRows(lRow - 1) + 1

	    Dim i
		Dim lngRangeFrom
		Dim lngRangeTo
		Dim strFlag
		Dim k

		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()

		k = 0
		totalQuotaRate = 0
		totalApportionQty     = 0

		for i = lngRangeFrom To lngRangeTo
			k = k + 1
			.vspdData2.Row = i
			.vspdData2.Col = 0
			strFlag = .vspdData2.Text

			if strFlag <> ggoSpread.DeleteFlag then
			    .vspdData2.Col = C_Quota_Rate
			    totalQuotaRate = totalQuotaRate + Unicdbl(.vspdData2.Text)
			    .vspdData2.Col = C_ApportionQty
			    totalApportionQty     = totalApportionQty     + Unicdbl(.vspdData2.Text)
		    end if
		Next

		iQuotaRate    = 100 - totalQuotaRate
		iApportionQty = iReqQty - totalApportionQty

        if iQuotaRate < 0 then iQuotaRate = 0
        if iApportionQty < 0 then iApportionQty = 0

		.vspdData2.Row = lRow2
		.vspdData2.Col = C_Quota_Rate
		.vspdData2.Text = UNIFormatNumber(iQuotaRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

		.vspdData2.Col = C_ApportionQty
    	.vspdData2.Text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)

		.vspdData.Row = lRow
		.vspdData.Col = C_CfmFlg
		.vspdData.value = 1

		.vspdData2.ReDraw = True
		.vspdData2.Col = C_SpplCd
		.vspdData2.focus
	End With
	Set gActiveElement = document.activeElement
	FncCopy = true
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
	FncCancel = false
	Dim lRow
	Dim i,k,iCnt
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim iActiveRow
	Dim iConvActiveRow
	Dim strFlag

	iActiveRow = frm1.vspdData.ActiveRow
	frm1.vspdData.Row = iActiveRow
	frm1.vspdData.Col = frm1.vspdData.MaxCols
	iConvActiveRow = frm1.vspdData.Text

	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End If

	If gActiveSpdSheet.ID = "B" Then

		'Check Spread2 Data Exists for the keys
		If CheckDataExist = False Then
			Exit function
		End If

		ggoSpread.Source = frm1.vspdData2
		With frm1.vspdData2
			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()

			.Redraw = False
			ggoSpread.EditUndo                                                  '��: Protect system from crashing
			Call checkdt(.ActiveRow)
			lngRangeFrom = ShowDataFirstRow2()
			lngRangeTo = ShowDataLastRow2()
			If lngRangeFrom > 0 Then
				iCnt=1
				For k=lngRangeFrom To lngRangeTo
					.Row=k
					.col=0
					if Isnumeric(.text) or Trim(.text)="" Then .text=iCnt
					iCnt = iCnt + 1
				Next
			End If
			.Redraw = True
		End With
	Else

		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo                                                  '��: Protect system from crashing

		if frm1.vspdData2.maxrowS > 0 Then
			ggoSpread.Source = frm1.vspdData2
			With frm1.vspdData2
				.Redraw = False

				lngRangeFrom = ShowDataFirstRow2()
				lngRangeTo = ShowDataLastRow2()

				If lngRangeFrom > 0 Then
					For k=lngRangeFrom to lngRangeTo
						.Row=k
						ggoSpread.EditUndo k                                                 '��: Protect system from crashing
						Call checkdt(k)
					Next
					lngRangeFrom = ShowDataFirstRow2()
					lngRangeTo = ShowDataLastRow2()
					iCnt=1
					For k=lngRangeFrom To lngRangeTo
						.Row=k
						.col=0
						if Isnumeric(.text) or Trim(.text)="" Then .text=iCnt
						iCnt = iCnt + 1
					Next
				End If
				.Redraw = True
			End WIth
		End If
	End If

	lRow = iActiveRow
	lngRangeFrom = ShowDataFirstRow2()
	lngRangeTo = ShowDataLastRow2()
	If lngRangeTo = 0 Then
		lglngHiddenRows(lRow - 1) = 0
	Else

		lglngHiddenRows(lRow - 1) = CInt(lngRangeTo) - CInt(lngRangeFrom) + 1
	End If

	k = 0
	If lngRangeFrom > 0 Then
		for i = lngRangeFrom to lngRangeTo
		    frm1.vspdData2.Row = i
		    frm1.vspdData2.Col = 0
		    strFlag = Trim(frm1.vspdData2.Text)
		    If strFlag = ggoSpread.InsertFlag or strFlag = ggoSpread.UpdateFlag or strFlag = ggoSpread.DeleteFlag then
		        k = 1
		        Exit for
		    End If
		next
	End If

	if k = 0 then
	    frm1.vspdData.Row = lRow
	    frm1.vspdData.Col = C_CfmFlg
	    frm1.vspdData.value = 0
	End If

	Set gActiveElement = document.activeElement
	FncCancel = true
End Function


'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)	'###�׸��� ������ ���Ǻκ�###
	FncInsertRow = false

	On Error Resume Next

	Dim lRow
	Dim lRow2
	Dim lconvRow
	Dim strMark
	Dim iInsertRow
	Dim IntRetCD
	Dim imRow
	Dim strInspUnitIndctnCd
	Dim iparentprno,iparentrow
	Dim totalQuotaRate,totalApportionQty
	Dim iQuotaRate,iApportionQty,iReqQty
	Dim iStrOrgCd

	With frm1
		If .vspdData.MaxRows <= 0 Then
			Exit Function
		End If

		.vspdData2.ReDraw = False

		If IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		Else
			imRow = AskSpdSheetAddRowCount()
			If imRow = "" Then
				Exit Function
			End If
		End If

		.vspdData2.focus
		ggoSpread.Source = .vspdData2
		ggoSpread.InsertRow .vspdData2.ActiveRow , imRow

		lRow = .vspdData.ActiveRow
		.vspdData.Row = lRow
		.vspdData.Col = .vspdData.MaxCols
		lconvRow = CInt(.vspdData.Text)

		.vspdData.Col = C_OrgCd
		iStrOrgCd = .vspdData.value

        .vspdData.Col = C_ReqNo
        iparentprno = .vspdData.value

        .vspdData.Col = C_ReqQty
        iReqQty = Unicdbl(.vspdData.text)

		For iInsertRow = 0 To imRow - 1
			lRow2 = .vspdData2.ActiveRow + iInsertRow

			.vspdData2.Row = lRow2
			.vspdData2.Col = 0
			strMark = .vspdData2.Text

			.vspdData2.Col = C_Flag
			.vspdData2.Text = strMark

			.vspdData2.Col = C_ParentRowNo
			.vspdData2.Text = lconvRow

			.vspdData2.Col = C_ParentPrNo
			.vspdData2.value = iparentprno

			.vspdData2.Col = C_OrgCd2
			.vspdData2.value = iStrOrgCd


			'�������� ���� �ش� Ű�� ���� Client�� Data Row���� ������ 
			lglngHiddenRows(lconvRow - 1) = CInt(lglngHiddenRows(lconvRow - 1)) + 1
			Call SetSpreadColor3(lRow2, lRow2)
		Next

		'/* ���� : ����� �� �ѹ��� ���� �߰� START */
		Dim i
		Dim lngRangeFrom
		Dim lngRangeTo
		Dim strFlag
		Dim k

		ggoSpread.Source = .vspdData2

		lngRangeFrom = ShowDataFirstRow2()
		lngRangeTo = ShowDataLastRow2()
		k = 0
		totalQuotaRate = 0
		totalApportionQty = 0

		for i = lngRangeFrom To lngRangeTo
			k = k + 1
			.vspdData2.Row = i
			.vspdData2.Col = 0
			strFlag = .vspdData2.Text
			If strFlag <> ggoSpread.InsertFlag and strFlag <> ggoSpread.UpdateFlag and strFlag <> ggoSpread.DeleteFlag then
				.vspdData2.Text = CStr(k)
			End If

			if strFlag <> ggoSpread.DeleteFlag then
			    .vspdData2.Col = C_Quota_Rate
			    totalQuotaRate = totalQuotaRate + Unicdbl(.vspdData2.Text)
			    .vspdData2.Col = C_ApportionQty
			    totalApportionQty     = totalApportionQty     + Unicdbl(.vspdData2.Text)
		    end if
		Next

		iQuotaRate = 100 - totalQuotaRate
		iApportionQty     = iReqQty - totalApportionQty

        if iQuotaRate < 0 then iQuotaRate = 0
        if iApportionQty < 0 then iApportionQty = 0

		.vspdData2.Row = lRow2
		.vspdData2.Col = C_Quota_Rate
		.vspdData2.Text = UNIFormatNumber(iQuotaRate, ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

		.vspdData2.Col = C_ApportionQty
    	.vspdData2.Text = UNIFormatNumber(iApportionQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	End With

	With frm1.vspdData
		.Row = lRow
		.Col = C_CfmFlg
		.value = 1
		.Col = 0
		.text = ggoSpread.UpdateFlag
	End With

	.vspdData2.ReDraw = True

	FncInsertRow = true

	Call SetSpreadLock()
	call chg_ass_method()
	Set gActiveElement = document.activeElement

End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()
	FncDeleteRow = false

	Dim lDelRows
	Dim iDelRowCnt, i
    Dim lngRangeFrom
    Dim lngRangeTo
    Dim iparentrow

	If frm1.vspdData.MaxRows < 1 then
		Exit function
	End if

	'Check Spread2 Data Exists for the keys
	If CheckDataExist = False Then
		Exit function
	End If

	With frm1.vspdData2
		.Redraw = False

		.Focus

		'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - START
	    lngRangeFrom = .SelBlockRow
	    .Row = lngRangeFrom
		If .RowHidden = True Then
			lngRangeFrom = ShowDataFirstRow2()
		End If

		lngRangeTo = .SelBlockRow2
		.Row = lngRangeTo
		If .RowHidden = True Then
			lngRangeTo = ShowDataLastRow2()
		End If

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Action = 2			'Select Block	SS_ACTION_SELECT_BLOCK
		.BlockMode = False
		'������ ������ �ʴ� ������ �Ѿ�� ��쿡 ���� ó�� - END

	    ggoSpread.Source = frm1.vspdData2
	     '----------  Coding part  -------------------------------------------------------------
		lDelRows = ggoSpread.DeleteRow
		.Row = lngRangeFrom
		.Col = C_ParentRowNo
		iparentrow = .text

		.BlockMode = True
		.Row = lngRangeFrom
		.Row2 = lngRangeTo
		.Col = 0
		.Col2 = 0
		.DestCol = C_Flag
		.DestRow = .SelBlockRow
		.Action = 19	'SS_ACTION_COPY_RANGE
		.BlockMode = False

		.Redraw = True
	End With

	With frm1.vspdData
		.Row = iparentrow
		.Col = C_CfmFlg
		.value = 1
		.Col = 0
		.text = ggoSpread.DeleteFlag
	End With
	Set gActiveElement = document.activeElement
	FncDeleteRow = true

	call chg_ass_method()
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
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function DbDelete()
	DbDelete = False
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()												        '���� ������ ���� ���� 
	DbDeleteOk = False
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

	 If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd=" & .hdnPlant.value
	    strVal = strVal & "&txtItemCd=" & .hdnItem.value
		strVal = strVal & "&txtFrDt=" & .hdnFrDt.Value
		strVal = strVal & "&txtToDt=" & .hdnToDt.value
		strVal = strVal & "&txtPoFrDt=" & .hdnPoFrDt.Value
		strVal = strVal & "&txtPoToDt=" & .hdnPoToDt.value
	    strVal = strVal & "&txtMRP=" & .hdnMrp.value
	    strVal = strVal & "&txtSoNo=" & .hdnSoNo.value
		strVal = strVal & "&txtTrkNo=" & .hdnTrkNo.value
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		strVal = strVal & "&rdoAppflg=" & .hdnFlg.value
	Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
	    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.Text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
	    strVal = strVal & "&txtPoFrDt=" & Trim(.txtPoFrDt.Text)
		strVal = strVal & "&txtPoToDt=" & Trim(.txtPoToDt.Text)
	    strVal = strVal & "&txtMRP=" & Trim(.txtMRP.value)
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    if .rdoAppflg(0).checked = true then
			strVal = strVal & "&rdoAppflg=" & "A"
		elseif .rdoAppflg(1).checked = true then
			strVal = strVal & "&rdoAppflg=" & "Y"
		else
			strVal = strVal & "&rdoAppflg=" & "N"
		End if
	    strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
		strVal = strVal & "&txtTrkNo=" & Trim(.txtTrackNo.value)
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

	With frm1
		lRow = .vspdData.MaxRows
		i=0
		If lRow > 0 And intARow > 0 Then
			frm1.btnAutoSel.disabled = False
			frm1.btnSelect.disabled = False
			frm1.btnDisSelect.disabled = False
			Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
			Call SetToolBar("11001111001011")				'��ư ���� ���� 

			If intTRow<=0 Then
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
			Else
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) ù��° �׸����� Ư��Row�� �ش��ϴ� �ι�° �׸����� Row ������ �����ϴ� �迭.
				For i = 0 To intTRow-1
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next
			End If

			For i = intTRow To intTRow+intARow-1
				lglngHiddenRows(i) = 0
			Next

			if lgIntFlgModeM = Parent.OPMD_CMODE then
			    If DbQuery2(1, False) = False Then	Exit Function
		    end if
		    lgIntFlgModeM = Parent.OPMD_UMODE
		    call chg_ass_method()
		Else
			frm1.btnAutoSel.disabled = true
			frm1.btnSelect.disabled = true
			frm1.btnDisSelect.disabled = True
			Call SetToolBar("11100000000011")				'��ư ���� ���� 
		End If
	End With
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
	    frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

    DbQueryOk = true
End Function

sub chg_ass_method()
	If frm1.rdoflg(0).checked = true then
		frm1.btnAutoSel.disabled = true
		frm1.btnSelect.disabled = true
		frm1.btnDisSelect.disabled = True
		frm1.vspdData.ReDraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLock C_CfmFlg , -1, C_CfmFlg, -1
		ggoSpread.SSSetProtected C_CfmFlg,			-1, -1
		frm1.vspdData.ReDraw = True
		Call SetToolBar("11001111001011")				'��ư ���� ���� 
	else
		frm1.btnAutoSel.disabled = false
		frm1.btnSelect.disabled = false
		frm1.btnDisSelect.disabled = false
		frm1.vspdData.ReDraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadunLock C_CfmFlg , -1, C_CfmFlg, -1
		'ggoSpread.SSSetRequired C_CfmFlg,			pvStartRow, pvEndRow
		frm1.vspdData.ReDraw = True
		Call SetToolBar("11000000000011")				'��ư ���� ���� 
	end if
end sub

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
    DbSave = False                                                          '��: Processing is NG
    Dim lRow
    Dim lRow2
	Dim lGrpCnt
	Dim strVal
	Dim strDel
	Dim iParentNum
	Dim iTargetParentNum
	Dim lngRangeFrom
    Dim lngRangeTo
    Dim parentRow
    Dim iReqQty
    Dim totalQty,totalRate
	Dim Zsep
	Dim iStrOrgCd
	Dim lngUpdCnt
	Dim iDelCnt
	Dim tmpRate, tmpQty
	Dim iColSep
	Dim iRowSep

	Dim strCUTotalvalLen '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�]
	Dim strDTotalvalLen  '���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]

	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 

	Dim iTmpCUBuffer         '������ ���� [����,�ű�]
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size

	Dim iTmpDBuffer          '������ ���� [����]
	Dim iTmpDBufferCount     '������ ���� Position
	Dim iTmpDBufferMaxCount  '������ ���� Chunk Size

	Call LayerShowHide(1)

	With frm1
		.txtMode.value = Parent.UID_M0002
	End With

	'-----------------------
	'Data manipulate area
	'-----------------------
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep
	lGrpCnt = 1
	strVal = ""
    strDel = ""
    Zsep = "@"
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0
	'-----------------------
	'Data manipulate area
	'-----------------------
	With frm1
		.vspdData.Col = C_ORGCd
	    iStrOrgCd = Trim(.vspdData.Text)
	    For parentRow = 1 To .vspdData.MaxRows

			If Trim(GetSpreadValue(.vspdData,C_CfmFlg,parentRow,"X","X")) = 1 Then

			    lngRangeFrom = DataFirstRow(parentRow)
			    lngRangeTo   = DataLastRow(parentRow)
			    iReqQty = Unicdbl(GetSpreadText(.vspdData,C_ReqQty,parentRow,"X","X"))
			    totalQty  = 0
			    totalRate = 0
			    lngUpdCnt = 0
			    iDelCnt = 0

			    For lRow = lngRangeFrom To lngRangeTo

					tmpQty = Unicdbl(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))
					tmpRate = Unicdbl(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))

					IF (tmpQty = 0 OR tmpRate = 0) AND  Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) <> ggoSpread.DeleteFlag Then
						Call DisplayMsgBox("970021","X","��к���(%),��з�","X")
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
					    Exit Function
					End if

					if Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) <> ggoSpread.DeleteFlag then
						'��з��� 
						totalRate = totalRate + Unicdbl(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))
						'��м��� �� 
						totalQty = totalQty + Unicdbl(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))
					end if

					If Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) = ggoSpread.InsertFlag OR Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) = ggoSpread.UpdateFlag OR Trim(GetSpreadText(.vspdData2,0,lRow,"X","X")) = ggoSpread.DeleteFlag Then
						lngUpdCnt = lngUpdCnt + 1
					End IF

					Select Case Trim(GetSpreadText(.vspdData2,0,lRow,"X","X"))

						Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

						    if Trim(GetSpreadText(.vspdData2,0,lRow,"X","X"))=ggoSpread.InsertFlag then
								strVal = strVal & "C" & iColSep
							Else
								strVal = strVal & "U" & iColSep
							End if

                            strVal = strVal & Trim(GetSpreadText(.vspdData2,C_SpplCd,lRow,"X","X")) & iColSep

						    If Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X"))="" Then
								strVal = strVal & "0" & iColSep
							Else
								strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_Quota_Rate,lRow,"X","X")),0) & iColSep
							End If

							If Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X"))="" Then
								strVal = strVal & "0" & iColSep
							Else
								strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData2,C_ApportionQty,lRow,"X","X")),0) & iColSep
							End If

							strVal = strVal & UNIConvDate(Trim(GetSpreadText(.vspdData2,C_PlanDt,lRow,"X","X"))) & iColSep
							strVal = strVal & Trim("" & GetSpreadText(.vspdData2,C_OrgCd2,lRow,"X","X")) & iColSep
							strVal = strVal & Trim("" & GetSpreadText(.vspdData2,C_GrpCd,lRow,"X","X")) & iColSep
							strVal = strVal & Trim("" & GetSpreadText(.vspdData2,C_ParentPrNo,lRow,"X","X")) & iColSep
							strVal = strVal & lngUpdCnt & iRowSep

							lGrpCnt = lGrpCnt + 1
						Case ggoSpread.DeleteFlag				'��: ���� 

							strDel = strDel & "D" & iColSep			'��: D=Delete
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
							strDel = strDel & Trim("" & GetSpreadText(.vspdData2,C_OrgCd2,lRow,"X","X")) & iColSep
							strDel = strDel & Trim("" & GetSpreadText(.vspdData2,C_GrpCd,lRow,"X","X")) & iColSep
							strDel = strDel & Trim("" & GetSpreadText(.vspdData2,C_ParentPrNo,lRow,"X","X")) & iColSep
							strDel = strDel & lngUpdCnt & iRowSep

                            iDelCnt = iDelCnt + 1
							lGrpCnt = lGrpCnt + 1

					End Select
			    Next

				If Trim(strVal) = "" and  Trim(strDel) = "" then
					strVal	=	""
				Else
					strVal = strDel & strVal & Zsep
				End iF

			    If iDelCnt-1 <> lngRangeTo - lngRangeFrom Then
					If totalQty <> iReqQty Then
					    Call DisplayMsgBox("172420", "X","{" & parentRow & "} Row", "X")
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
					    Exit Function
					End If

					If totalRate <> 100 Then
					    Call DisplayMsgBox("171325", "X", "X", "X")
					    Call LayerShowHide(0)
					    Call RemovedivTextArea
					    Exit Function
					End If
				End if

				Select Case Trim(GetSpreadText(.vspdData,0,parentRow,"X","X"))
				    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag,ggoSpread.DeleteFlag
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
				End Select
			End If
			strVal  = ""
			strDel  = ""
		Next

	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)			'��: �����Ͻ� ASP �� ���� 

	DbSave = True
End Function


'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	Call InitVariables
	frm1.vspdData2.MaxRows = 0
	Call MainQuery()
End Function

'========================================== btnAuto()  =============================================
'	Name : btnAuto()
'	Description : ��ü�ڵ����� ��ư�� Click ���� ��� 
'=========================================================================================================
Function btnAuto()
   Dim IntRetCD
   Dim IntChkCnt
   Dim parentRow

   btnAuto = False
   IntChkCnt = 0

   For parentRow = 1 To frm1.vspdData.MaxRows
	     frm1.vspdData.Row = parentRow
	     frm1.vspdData.Col = C_CfmFlg

		 If Trim(frm1.vspdData.value) = 1 then
			IntChkCnt = IntChkCnt + 1
		 End If
   Next

	If IntChkCnt = 0 Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
	    Exit Function
	End IF

   IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "X", "X")
   If IntRetCD = vbNo Then Exit Function

   Err.Clear

    '-----------------------
    'Save function call area
    '-----------------------
   Call DbAutoSave

   btnAuto = True
End Function
'========================================== DbAutoSave()  =============================================
'	Name : DbAutoSave()
'	Description : ��ü�ڵ����� ��ư�� Click ���� ��� DbSave��� DbAutoSave�Լ��� ȣ���Ѵ�.
'=========================================================================================================
Function DbAutoSave()
    Dim lRow
    Dim lGrpCnt
	Dim strVal
	Dim parentRow
	Dim igColSep,igRowSep
    DbAutoSave = False

	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep

    If LayerShowHide(1) = False Then Exit Function

	With frm1
		.txtMode.value = "AutoAssign"

		lGrpCnt = 1

		strVal = ""

		If .rdoAssflg(0).checked = true then
		    .hdnrdoAssflg.value = "D"
		Elseif .rdoAssflg(1).checked = true then
			.hdnrdoAssflg.value = "R"
		Else
			.hdnrdoAssflg.value = "Q"
		End if

		'-----------------------
		'Data manipulate area
		'-----------------------
		For parentRow = 1 To .vspdData.MaxRows

			.vspdData.Row = parentRow
	        .vspdData.Col = C_CfmFlg

			if Trim(.vspdData.value) = 1 then

				strVal = strVal & .hdnrdoAssflg.value & igColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_ReqNo,parentRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,parentRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemCd,parentRow,"X","X")) & igColSep
				strVal = strVal & parentRow & igRowSep

				lGrpCnt = lGrpCnt + 1
			end if
		next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		if Trim(strVal) <> "" then
			If LayerShowHide(1) = False Then Exit Function
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		End if

	End With
    DbAutoSave = True                                                           '��: Processing is NG

End Function

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : ���ó����� ��ư�� Ŭ���� ��� �߻� 
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
'   Event Desc : ���ó����� ��ư�� Ŭ���� ��� �߻� 
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

Sub chkAssignflag()
'//�ּ�ó����:2005-09-12
'	With frm1.vspdData
'		If .MaxRows <= 0 Then Exit Sub
'		.Row = .ActiveRow
'	    .Col = C_CfmFlg
'		.value = 1
'	End With
'	Set gActiveSpdSheet = frm1.vspdData
End Sub
'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 																		#
'########################################################################################################
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ü����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm" SIZE=25 tag="14X"></TD>

								<TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemcd" SIZE=15 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">
													   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtitemNm" SIZE=25 tag="14X"></TD>
							</TR>

							<TR>
								<TD CLASS="TD5" NOWRAP>��ü��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="��ü��������" NAME="rdoAppflg" id = "rdoAppflg1" Value="A" checked tag="11"><label for="rdoAppflg1">&nbsp;��ü&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="��ü��������" NAME="rdoAppflg" id = "rdoAppflg2" Value="Y" tag="11"><label for="rdoAppflg2">&nbsp;����&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="��ü��������" NAME="rdoAppflg" id = "rdoAppflg3" Value="N" tag="11"><label for="rdoAppflg3">&nbsp;������&nbsp;</label></TD>
							    <TD CLASS="TD5" NOWRAP>MRP Run��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="MRP Run��ȣ" NAME="txtMRP" SIZE=26 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMrp"></TD>
							</TR>


							<TR><TD CLASS="TD5" NOWRAP>�ʿ���</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
											<td>~</td>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ʿ���" NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
										<tr>
									</table>
								</TD>
								<TD CLASS="TD5" NOWRAP>���ֿ�����</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="���ֿ�����" NAME="txtPoFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
											<td>~</td>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="���ֿ�����" NAME="txtPoToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="���ֹ�ȣ" NAME="txtSoNo" SIZE=20 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo"></TD>
								<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="Tracking No." NAME="txtTrackNo" SIZE=26 MAXLENGTH=25 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��ü�������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="��ü�������" NAME="rdoflg" id = "rdoflg1" Value="M" checked tag="11" onClick=""><label for="rdoflg1">&nbsp;����&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="��ü�������" NAME="rdoflg" id = "rdoflg2" Value="A"         tag="11" onClick=""><label for="rdoflg2">&nbsp;�ڵ�&nbsp;</label></TD>
								<TD CLASS="TD5" NOWRAP>�ڵ��������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="�ڵ��������" NAME="rdoAssflg" id = "rdoAssflg1" Value="A" checked tag="11" onClick="chkAssignflag()"><label for="rdoAssflg1">&nbsp;�ְ��޼�&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="�ڵ��������" NAME="rdoAssflg" id = "rdoAssflg2" Value="Y" tag="11" onClick="chkAssignflag()"><label for="rdoAssflg2">&nbsp;RULE&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="�ڵ��������" NAME="rdoAssflg" id = "rdoAssflg3" Value="N" tag="11" onClick="chkAssignflag()"><label for="rdoAssflg3">&nbsp;��к�&nbsp;</label></TD>
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>

    <TR HEIGHT= 40%>
     <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
     <!-- <TR>
		<TD CLASS="TD5" NOWRAP>��ü�������</TD>
		<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="��ü�������" NAME="rdoflg" id = "rdoflg1" Value="M" checked tag="11" onClick="chg_ass_method()"><label for="rdoflg1">&nbsp;����&nbsp;</label>
							   <INPUT TYPE=radio Class="Radio" ALT="��ü�������" NAME="rdoflg" id = "rdoflg2" Value="A"         tag="11" onClick="chg_ass_method()"><label for="rdoflg2">&nbsp;�ڵ�&nbsp;</label></TD>
		<TD CLASS="TD5" NOWRAP>�ڵ��������</TD>
		<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="�ڵ��������" NAME="rdoAssflg" id = "rdoAssflg1" Value="A" checked tag="11" onClick="chkAssignflag()"><label for="rdoAssflg1">&nbsp;�ְ��޼�&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="�ڵ��������" NAME="rdoAssflg" id = "rdoAssflg2" Value="Y" tag="11" onClick="chkAssignflag()"><label for="rdoAssflg2">&nbsp;RULE&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="�ڵ��������" NAME="rdoAssflg" id = "rdoAssflg3" Value="N" tag="11" onClick="chkAssignflag()"><label for="rdoAssflg3">&nbsp;��к�&nbsp;</label></TD>
	  </TR> -->
       <TR>
        <TD HEIGHT=100% WIDTH=100% COLSPAN=4>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="A"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>

    <TR>
     <TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
    </TR>

    <TR HEIGHT= 40%>
     <TD WIDTH=100% HEIGHT=* VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD HEIGHT=100% WIDTH=100% COLSPAN=4>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id="B"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>

   </TABLE>
  </TD>
 </TR>
    <tr>
  <TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
  <td WIDTH="100%">
   <table <%=LR_SPACE_TYPE_30%>>
    <tr>
     <TD WIDTH=10>&nbsp;</TD>
     <td WIDTH="*" align="left">
     <button name="btnAutoSel" class="clsmbtn" ONCLICK="btnAuto()">��ü�ڵ�����</button>&nbsp;
     <BUTTON NAME="btnSelect" CLASS="CLSMBTN">�ϰ�����</BUTTON>
     <BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">�ϰ��������</BUTTON>
     </td>
     <TD WIDTH=10>&nbsp;</TD>
    </tr>
   </table>
  </td>
    </tr>
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex = -1></IFRAME>
  </TD>
 </TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" tabindex = -1></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnMrp" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnflg" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnSoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrkNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoFrDt" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnPoToDt" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnrdoAssflg" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnState" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnOrg" tag="24" tabindex = -1>
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
