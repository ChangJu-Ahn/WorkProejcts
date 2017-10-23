<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Change
'*  3. Program ID           : a7127ma1
'*  4. Program Name         : �Ű�/��⳻�� �ϰ����
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/03/01
'*  8. Modified date(Last)  : 2003/03/20
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. �� �� ��
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'=======================================================================================================
'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit    		'��: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ ��
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>
'@PGM_ID
Const BIZ_PGM_ID  = "a7127mb1.asp"  
Const BIZ_LOAD_ID  = "a7127mb2.asp"  
											'�����Ͻ� ���� ASP��
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'ȯ������ �����Ͻ� ���� ASP��

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

'@Grid_Column
Dim	C_ChgNo
Dim	C_AsstNo
Dim	C_AsstNm
Dim C_SubNo
Dim	C_DeptCd
Dim	C_DeptNm
Dim C_OrgChgId
Dim	C_AcqDt
Dim	C_InvQty
Dim	C_ChgQty
Dim C_SoldRate
Dim	C_ChgAmt
Dim	C_ChgLocAmt
Dim	C_AcqLocAmt
Dim	C_DeprLocAmt
Dim	C_BALLocAmt
Dim	C_MnthDeprAmt
Dim	C_TaxAmt
Dim	C_TaxLocAmt
Dim	C_AccDeprAmt
Dim	C_AsstSoldDesc

Const C_SHEETMAXROWS = 30							            '�� ȭ�鿡 �������� �ִ밹��

'@Grid_Column
Dim C_RcptTypeCd
Dim C_RcptTypePopup
Dim C_RcptTypeNm							            'Spread Sheet �� Columns �ε���
Dim C_RcptAmt
Dim C_RcptLocAmt
Dim C_ARAPNo
Dim C_ArAcctCd
Dim C_ArAcctPopup
Dim C_ArAcctNm
Dim C_ArDueDt
Dim C_BankCd								            'Spread Sheet �� Columns �ε���
Dim C_BankPopup
Dim C_BankNm
Dim C_BankAcctCd
Dim C_NoteNo
Dim C_NotePopup
Dim C_RcptDesc

Const C_SHEETMAXROWS2 = 30							            '�� ȭ�鿡 �������� �ִ밹��

Dim lgStrPrevKey2
Dim IsOpenPop						                        'Popup
Dim gSelframeFlg                                            'Current Tab Page

Dim lgMasterQueryFg                                         ''�ڻ�Master�� query ����

' ���Ѱ��� �߰�
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' �����
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ����

'======================================================================================================
'												2. Function��
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

Sub initSpreadPosVariables()


'Grid 1 vspddata 
	C_ChgNo				= 1
	C_AsstNo			= 2
	C_AsstNm			= 3
	C_SubNo				= 4 
	C_DeptCd			= 5 
	C_DeptNm			= 6 
	C_OrgChgId			= 7 
	C_AcqDt				= 8 
	C_InvQty			= 9 
	C_ChgQty			= 10
	C_SoldRate			= 11
	C_ChgAmt			= 12
	C_ChgLocAmt			= 13
	C_AcqLocAmt			= 14
	C_DeprLocAmt		= 15
	C_BALLocAmt			= 16
	C_MnthDeprAmt		= 17
	C_TaxAmt			= 18
	C_TaxLocAmt			= 19
	C_AccDeprAmt		= 20
	C_AsstSoldDesc		= 21

'Grid 2 vspddata2

	C_RcptTypeCd		= 1
	C_RcptTypePopup		= 2
	C_RcptTypeNm		= 3						            'Spread Sheet �� Columns �ε���
	C_RcptAmt			= 4
	C_RcptLocAmt		= 5
	C_ARAPNo			= 6
	C_ArAcctCd			= 7
	C_ArAcctPopup		= 8
	C_ArAcctNm			= 9
	C_ArDueDt			= 10
	C_BankCd			= 11						            'Spread Sheet �� Columns �ε���
	C_BankPopup			= 12
	C_BankNm			= 13
	C_BankAcctCd		= 14
	C_NoteNo			= 15
	C_NotePopup			= 16
	C_RcptDesc			= 17

End Sub


'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                   'Indicates that current mode is Create mode
                                        'Indicates that no value changed
    lgIntGrpCount = 0                                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
	lgBlnFlgChgValue = False
    lgStrPrevKey = ""                                           'initializes Previous Key
	lgStrPrevKey2= ""
	
	gSelframeFlg = TAB1
	
	
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
<%
	Dim svrDate
	svrDate = GetSvrDate
%>

	if lgIntFlgMode = parent.OPMD_CMODE then
		frm1.txtChgDt.text		= UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
		frm1.txtIssuedDt.text	= UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	
	end if
	
	frm1.txtDocCur.value	= parent.gCurrency

	If gIsShowLocal <> "N" Then
		frm1.txtXchRate.text	= "1"
	else
		frm1.txtXchRate.value	= "1"	
	end if
	frm1.txtVatRate.text = "0"
	lgBlnFlgChgValue = False
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet(strval)
    Call InitSpreadPosVariables()
	Select Case UCase(strval)
	Case "A"
		With frm1.vspdData
		
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20030513",,parent.gAllowDragDropSpread  

			.ReDraw = false	

			.MaxCols = C_AsstSoldDesc + 1                               '��: �ִ� Columns�� �׻� 1�� ������Ŵ
			ggoSpread.Source = frm1.vspdData
			ggospread.ClearSpreadData		'Buffer Clear

			'Hidden Column ����
			.Col = .MaxCols											'������Ʈ�� ��� Hidden Column
			.ColHidden = True

			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_ChgNo,		"������ȣ",		16,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_AsstNo,		"�ڻ��ڵ�",		10,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_AsstNm,		"�ڻ��",		14,		0,		-1,		40
			ggoSpread.SSSetEdit		C_SubNo,		"Sub No",		8,		0,		-1,		18,		2
			'ggoSpread.SSSetFloat		C_SubNo,		"Sub No",		8,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetEdit		C_DeptCd,		"�μ��ڵ�",		10,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_OrgChgId,		"��������ID",	10,		0,		-1,		18,		2
			ggoSpread.SSSetEdit		C_DeptNm,		"�μ���",		14,		0,		-1,		40
			ggoSpread.SSSetDate		C_AcqDt,		"�������",		10,		2,		gDateFormat  			
			Call AppendNumberPlace("6","11","0")
			ggoSpread.SSSetFloat    C_InvQty,		"������",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat    C_ChgQty,		"��������",		12,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,, "Z", "1", "100000"

			ggoSpread.SSSetFloat	C_SoldRate,		"�Ű�����(%)",	12, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			
			ggoSpread.SSSetFloat	C_ChgAmt,		"�Ǹž�",		12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_ChgLocAmt,	"�Ǹž�(�ڱ�)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_AcqLocAmt,	"�������ݾ�(�ڱ�)",		15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_DeprLocAmt,	"���һ󰢴����(�ڱ�)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_BALLocAmt,	"�ڻ꺯����(�ڱ�)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_MnthDeprAmt,	"�����󰢴���(�ڱ�)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_TaxAmt,		"�ΰ����ݾ�",		12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_TaxLocAmt,	"�ΰ����ݾ�(�ڱ�)",	15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_AccDeprAmt,	"�ڻ갨�Ҿ�",	14, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"

			ggoSpread.SSSetEdit		C_AsstSoldDesc,	"����",				20,		0,		-1,		40

			Call ggoSpread.SSSetColHidden(C_ChgNo,C_ChgNo,True)
			Call ggoSpread.SSSetColHidden(C_OrgChgId,C_OrgChgId,True)
			Call ggoSpread.SSSetColHidden(C_MnthDeprAmt,C_MnthDeprAmt,True)
			Call ggoSpread.SSSetColHidden(C_AccDeprAmt,C_AccDeprAmt,True)

			.ReDraw = true
		
		End With
		Call SetSpreadLock("A")
		
	Case "B"
	
		With frm1.vspdData2
		
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20030513",,parent.gAllowDragDropSpread  
			.ReDraw = false	
		
			.MaxCols = C_RcptDesc + 1                               '��: �ִ� Columns�� �׻� 1�� ������Ŵ
			ggoSpread.Source = frm1.vspdData2
			ggospread.ClearSpreadData		'Buffer Clear

			'Hidden Column ����
			.Col = .MaxCols											'������Ʈ�� ��� Hidden Column
			.ColHidden = True
				
			Call GetSpreadColumnPos("B")
			
			ggoSpread.SSSetEdit		C_RcptTypeCd,	"�Ա�����",	10,		0,		-1,		10,		2
			ggoSpread.SSSetButton	C_RcptTypePopup
			ggoSpread.SSSetEdit		C_RcptTypeNm,	"�Ա�������",	12,		0,		-1,		10,		2
			ggoSpread.SSSetFloat	C_RcptAmt,		"�ݾ�",			14, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetFloat	C_RcptLocAmt,	"�ݾ�(�ڱ�)",	14, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"0"
			ggoSpread.SSSetEdit		C_ARAPNo,		"�̼��ݹ�ȣ",		10,		0,		-1,		30
			ggoSpread.SSSetEdit		C_ArAcctCd,		"�̼��ݰ����ڵ�",	14,		0,		-1,		18,		2
			ggoSpread.SSSetButton	C_ArAcctPopup
			ggoSpread.SSSetEdit		C_ArAcctNm,		"�̼��ݰ�����",		14,		0,		-1,		40,		2
			ggoSpread.SSSetDate		C_ArDueDt,		"�̼��ݸ�������",	14,		2,		gDateFormat  
			ggoSpread.SSSetEdit		C_BankCd,		"�����ڵ�",	10,		0,		-1,		10,		2
			ggoSpread.SSSetButton	C_BankPopup
			ggoSpread.SSSetEdit		C_BankNm,		"�����",		14,		0,		-1,		40
			ggoSpread.SSSetEdit		C_BankAcctCd,	"���¹�ȣ",		12,		0,		-1,		30,		2
			ggoSpread.SSSetEdit		C_NoteNo,		"������ȣ",		10,		0,		-1,		30,		2
			ggoSpread.SSSetButton	C_NotePopup			
			ggoSpread.SSSetEdit		C_RcptDesc,		"����",				20,		0,		-1,		40

			Call ggoSpread.MakePairsColumn(C_RcptTypeCd,C_RcptTypePopup)
			Call ggoSpread.MakePairsColumn(C_ArAcctCd,C_ArAcctPopup)
			Call ggoSpread.MakePairsColumn(C_BankCd,C_BankPopup)
			Call ggoSpread.MakePairsColumn(C_NoteNo,C_NotePopup)

			.ReDraw = true

		End With
		Call SetSpreadLock("B")
			
	End Select
End Sub


'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock(strval)
    With frm1
		Select Case UCase(strval)
		Case "A"
			.vspdData.ReDraw = False
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLock	C_ChgNo,			-1,		C_ChgNo
			ggoSpread.SpreadLock	C_AsstNo,			-1,		C_AsstNo
			ggoSpread.SpreadLock	C_AsstNm,			-1,		C_AsstNm
			ggoSpread.SpreadLock	C_SubNo,			-1,		C_SubNo			
			ggoSpread.SpreadLock	C_DeptCd,			-1,		C_DeptCd
			ggoSpread.SpreadLock	C_DeptNm,			-1,		C_DeptNm
			ggoSpread.SpreadLock	C_OrgChgId,			-1,		C_OrgChgId
			ggoSpread.SpreadLock	C_AcqDt,			-1,		C_AcqDt
			ggoSpread.SpreadLock	C_AcqLocAmt,		-1,		C_AcqLocAmt
			ggoSpread.SpreadLock	C_InvQty,			-1,		C_InvQty
			ggoSpread.SSSetRequired		C_SoldRate,		-1,		C_SoldRate
			ggoSpread.SpreadLock	C_DeprLocAmt,		-1,		C_DeprLocAmt
			ggoSpread.SpreadLock	C_BALLocAmt,		-1,		C_BALLocAmt
			ggoSpread.SpreadLock	C_MnthDeprAmt,		-1,		C_MnthDeprAmt
			ggoSpread.SSSetRequired	C_ChgQty,			-1,		C_ChgQty
			ggoSpread.SSSetRequired	C_ChgAmt,			-1,		C_ChgAmt
			ggoSpread.SSSetRequired	C_SoldRate,			-1,		C_SoldRate			
			ggoSpread.SpreadLock	C_AccDeprAmt,		-1,		C_AccDeprAmt
			
			.vspdData.ReDraw = True

		Case "B"
			.vspdData2.ReDraw = False
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SSSetRequired	C_RcptTypeCd,		-1,		C_RcptTypeCd
			ggoSpread.SpreadLock	C_RcptTypeNm,		-1,		C_RcptTypeNm
			ggoSpread.SSSetRequired	C_RcptAmt,			-1,		C_RcptDesc
			ggoSpread.SpreadLock	C_ARAPNo,			-1,		C_ARAPNo
			ggoSpread.SpreadLock	C_ArAcctCd,			-1,		C_ArAcctCd
			ggoSpread.SpreadLock	C_ArAcctPopup,		-1,		C_ArAcctPopup
			ggoSpread.SpreadLock	C_ArAcctNm,			-1,		C_ArAcctNm
			ggoSpread.SpreadLock	C_ArDueDt,			-1,		C_ArDueDt
			ggoSpread.SpreadLock	C_BankCd,			-1,		C_BankCd
			ggoSpread.SpreadLock	C_BankPopup,		-1,		C_BankPopup
			ggoSpread.SpreadLock	C_BankNm,			-1,		C_BankNm
			ggoSpread.SpreadLock	C_BankAcctCd,		-1,		C_BankAcctCd
			ggoSpread.SpreadLock	C_NoteNo,			-1,		C_NoteNo
			ggoSpread.SpreadLock	C_NotePopup,		-1,		C_NotePopup

			.vspdData2.ReDraw = True
		End Select
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow,byVal SpreadTab)
   	With frm1
	If Trim(SpreadTab) = "A" Then
		ggoSpread.Source = frm1.vspdData

		.vspdData.ReDraw = False

			ggoSpread.SSSetProtected	C_ChgNo,		pvStarRow, pvEndRow	
			ggoSpread.SSSetProtected	C_AsstNo,		pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_AsstNm,		pvStarRow, pvEndRow	
			ggoSpread.SSSetProtected	C_DeptCd,		pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_DeptNm,		pvStarRow, pvEndRow	
			ggoSpread.SSSetProtected	C_OrgChgId,	pvStarRow, pvEndRow	
			ggoSpread.SSSetProtected	C_AcqDt,		pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_AcqLocAmt,	pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_InvQty,		pvStarRow, pvEndRow
'			ggoSpread.SpreadLock		C_SoldRate,	pvStarRow, pvEndRow	
			ggoSpread.SSSetProtected	C_SoldRate,		pvStarRow, pvEndRow

			ggoSpread.SSSetProtected	C_DeprLocAmt,		pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_BALLocAmt,		pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_MnthDeprAmt,		pvStarRow, pvEndRow
			ggoSpread.SSSetRequired		C_ChgQty,		pvStarRow, pvEndRow
			ggoSpread.SSSetRequired		C_ChgAmt,		pvStarRow, pvEndRow
			ggoSpread.SSSetProtected	C_AccDeprAmt,	pvStarRow, pvEndRow

			If frm1.Rb_Duse.checked = true then 
				ggoSpread.SSSetProtected	C_ChgAmt,		1, frm1.vspdData.MaxRows
				ggoSpread.SSSetProtected	C_ChgLocAmt,	1, frm1.vspdData.MaxRows
				ggoSpread.SSSetProtected	C_AccDeprAmt,	1, frm1.vspdData.MaxRows
				ggoSpread.SSSetProtected	C_TaxAmt,		1, frm1.vspdData.MaxRows
				ggoSpread.SSSetProtected	C_TaxLocAmt,	1, frm1.vspdData.MaxRows
			End If
			
		.vspdData.ReDraw = True
	Else
		ggoSpread.Source = frm1.vspdData2

		.vspdData2.ReDraw = False

			ggoSpread.SSSetRequired	C_RcptTypeCd,	pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_RcptTypeNm,	pvStarRow, pvEndRow	
			ggoSpread.SSSetRequired	C_RcptAmt,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_ARAPNo,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_ArAcctCd,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_ArAcctPopup,	pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_ArAcctNm,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_ArDueDt,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_BankCd,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_BankPopup,	pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_BankNm,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_BankAcctCd,	pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_NoteNo,		pvStarRow, pvEndRow	
			ggoSpread.SpreadLock	C_NotePopup,	pvStarRow, pvEndRow	

		.vspdData2.ReDraw = True
	End If
	End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ChgNo				= iCurColumnPos(1)
			C_AsstNo			= iCurColumnPos(2)
			C_AsstNm			= iCurColumnPos(3)
			C_SubNo             = iCurColumnPos(4)
			C_DeptCd			= iCurColumnPos(5)
			C_DeptNm			= iCurColumnPos(6)
			C_OrgChgId			= iCurColumnPos(7)
			C_AcqDt				= iCurColumnPos(8)
			C_InvQty			= iCurColumnPos(9)
			C_ChgQty			= iCurColumnPos(10)
			C_SoldRate			= iCurColumnPos(11)
			C_ChgAmt			= iCurColumnPos(12)
			C_ChgLocAmt			= iCurColumnPos(13)
			C_AcqLocAmt			= iCurColumnPos(14)
			C_DeprLocAmt		= iCurColumnPos(15)
			C_BALLocAmt			= iCurColumnPos(16)
			C_MnthDeprAmt		= iCurColumnPos(17)
			C_TaxAmt			= iCurColumnPos(18)
			C_TaxLocAmt			= iCurColumnPos(19)
			C_AccDeprAmt		= iCurColumnPos(20)
			C_AsstSoldDesc		= iCurColumnPos(21)

	Case "B"
		ggoSpread.Source = frm1.vspdData2
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_RcptTypeCd	= iCurColumnPos(1)
			C_RcptTypePopup	= iCurColumnPos(2)
			C_RcptTypeNm	= iCurColumnPos(3)									            'Spread Sheet �� Columns �ε��� 
			C_RcptAmt		= iCurColumnPos(4)								            'Spread Sheet �� Columns �ε��� 
			C_RcptLocAmt	= iCurColumnPos(5)
			C_ARAPNo		= iCurColumnPos(6)
			C_ArAcctCd		= iCurColumnPos(7)
			C_ArAcctPopup	= iCurColumnPos(8)
			C_ArAcctNm		= iCurColumnPos(9)
			C_ArDueDt		= iCurColumnPos(10)
			C_BankCd		= iCurColumnPos(11)
			C_BankPopup		= iCurColumnPos(12)
			C_BankNm		= iCurColumnPos(13)
			C_BankAcctCd	= iCurColumnPos(14)
			C_NoteNo		= iCurColumnPos(15)
			C_NotePopup		= iCurColumnPos(16)
			C_RcptDesc		= iCurColumnPos(17)

	End Select
End Sub

 '==========================================  2.3.1 Tab Click ó��  =================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=================================================================================================================== 
 '----------------  ClickTab1(): Header Tabó�� �κ� (Header Tab�� �ִ� ��츸 ���)  ---------------------------- 
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB1
	
	If lgIntFlgMode = parent.OPMD_UMODE then
		Call SetToolBar("11111011000111")
	else 
		Call SetToolBar("11101001000111")
	end if
	
End Function

Function ClickTab2()
	
	If frm1.Rb_Duse.checked then Exit Function
	
	If gSelframeFlg = TAB2 Then Exit Function

	Call changeTabs(TAB2)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB2

	If lgIntFlgMode = parent.OPMD_UMODE then
		Call SetToolBar("11111111001111")
	else 
		Call SetToolBar("11101101001111")
	end if
	
End Function
 
'======================================================================================================
'   Function Name : OpenChgNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenChgNoInfo()
	Dim arrRet
	Dim IntRetCD
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("A7127RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A7127RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtAsstChgNo.focus
		Exit Function
	Else
		Call SetChgNoInfo(arrRet)
	End If	
 
End Function

'======================================================================================================
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetChgNoInfo(Byval arrRet)

	With frm1
		.txtAsstChgNo.value  = arrRet(0)			
		.txtAsstChgNo.focus
	End With

End Function

'=======================================================================================================
'	Name : OpenDeptCd()
'	Description : Dept Cd PopUp
'=======================================================================================================
Function OpenDeptCd(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("DeptPopupDtA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtChgDt.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  

	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

' T : protected F: �ʼ� 
	If lgIntFlgMode = parent.OPMD_UMODE then
		arrParam(3) = "T"									' �������� ���� Condition  
	Else
		If frm1.txtChgDt.className = parent.UCN_PROTECTED Then
			arrParam(3) = "T"									' �������� ���� Condition  
		Else
			arrParam(3) = "F"									' �������� ���� Condition  
		End If
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDeptCd(arrRet, iWhere)
	End If
	
End Function

'=======================================================================================================
'	Name : SetDeptCd()
'	Description : DeptCd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetDeptCd(byval arrRet, byval iWhere)
	frm1.txtDeptCd.Value    = arrRet(0)
	frm1.txtDeptNm.Value    = arrRet(1)

	If frm1.txtChgDt.className <> parent.UCN_PROTECTED Then
		frm1.txtChgDt.text		= arrRet(3)
	End If

	Call txtDeptCd_OnChange()

	frm1.txtDeptCd.focus
	
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
		lgBlnFlgChgValue = True
	End If
		
End Function
'========================================================================================
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.focus
	frm1.txtBpCd.Value    = Trim(arrRet(0))
	frm1.txtBpNm.Value    = Trim(arrRet(1))		
	lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(Byval strCode,Byval strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function
		
	IF UCase(strCard) = "CR"	Then		
		arrParam(0) = "���뱸��ī�� �˾�"				        ' �˾� ��Ī
		arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d"		' TABLE ��Ī
		arrParam(2) = ""								' Code Condition
		arrParam(3) = ""								' Name Cindition			
		arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CR", "''", "S") & "  and a.bp_cd = b.bp_cd  "			
		arrParam(4) = arrParam(4) & " and a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "
		arrParam(5) = "����ī���ȣ"						' �����ʵ��� �� ��Ī

		arrField(0) = "a.Note_no"					' Field��(0)
		arrField(1) = "F2" & parent.gColSep & "a.Note_amt"		' Field��(1)
		arrField(2) = "DD" & parent.gColSep & "a.Issue_dt"		' Field��(2)
		arrField(3) = "b.bp_nm"					' Field��(3)
		arrField(4) = "d.card_co_nm"    	    			' Field��(4)

		arrHeader(0) = "����ī���ȣ"				' Header��(0)
		arrHeader(1) = "�ݾ�"				' Header��(1)
		arrHeader(2) = "������"				' Header��(2)	    
		arrHeader(3) = "�ŷ�ó"				' Header��(3)
		arrHeader(4) = "ī���"				' Header��(4)

	Else

		arrParam(0) = "������ȣ �˾�"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
	
		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "������ȣ"			
	
		arrField(0) = "A.NOTE_NO"		
		arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
		arrField(2) = "C.BP_NM"	    
		arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
		arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
		arrField(5) = "B.BANK_NM"	        
    
		arrHeader(0) = "������ȣ"
		arrHeader(1) = "�����ݾ�"        		
		arrHeader(2) = "�ŷ�ó"        		        	
		arrHeader(3) = "������"        		        
		arrHeader(4) = "������"        		        
		arrHeader(5) = "����"
	End if
	        		        
	IsOpenPop = True	
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetNoteNo(arrRet)
	End If	
	
End Function

'=======================================================================================================
'	Name : SetNoteNo()
'	Description : Note No Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetNoteNo(byval arrRet)
	With frm1
		
		.vspdData2.Col	= C_NoteNo
		.vspdData2.Text	= arrRet(0)
			    
		.vspdData2.Col	= C_RcptLocAmt
		.vspdData2.Text	= arrRet(1)
		
	    Call vspdData2_Change(.vspdData2.Col, .vspdData2.Row)				 ' ������ dlf��ٰ� �˷��� 
		
		lgBlnFlgChgValue = True
	End With
End Function

'=======================================================================================================
'	Name : OpenCurrency()
'	Description : Currency PopUp
'=======================================================================================================
Function OpenCurrency()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg
    
    if frm1.Rb_Duse.checked = True then Exit Function
	If IsOpenPop = True Then Exit Function

	arrParam(0) = "�ŷ���ȭ �˾�"	
	arrParam(1) = "B_CURRENCY"				
	arrParam(2) = Trim(frm1.txtDocCur.Value)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "�ŷ���ȭ"
	
    arrField(0) = "CURRENCY"	
    arrField(1) = "CURRENCY_DESC"	
    
    arrHeader(0) = "�ŷ���ȭ"		
    arrHeader(1) = "�ŷ���ȭ��"

	IsOpenPop = True
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDocCur.focus
		Exit Function
	Else
		Call SetCurrency(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetCurrency()
'	Description : Currency Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCurrency(byval arrRet)
	frm1.txtDocCur.value    = arrRet(0)		
	
	If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' �ŷ���ȭ�ϰ� Company ��ȭ�� �ٸ��� ȯ���� 0���� ����
		If gIsShowLocal <> "N" Then
			frm1.txtXchRate.text	= "0"                   
		else
			frm1.txtXchRate.value	= "0" 								
		end if							                       							                                        
	Else 
		If gIsShowLocal <> "N" Then
			frm1.txtXchRate.text	= "1"        
		else
			frm1.txtXchRate.value	= "1" 								
		end if							         								
	End If		
    
    call txtDocCur_OnChange()
    
    frm1.txtDocCur.focus
    
    lgBlnFlgChgValue = True
	
End Function

'=======================================================================================================
'	Name : OpenBankAcct()
'	Description : Bank Account No PopUp
'=======================================================================================================
Function OpenBankAcct(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function	
	
	arrParam(0) = "�������ڵ� �˾�"	' �˾� ��Ī
	arrParam(1) = "B_BANK A, F_DPST B"			' TABLE ��Ī
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "A.BANK_CD = B.BANK_CD "		' Where Condition
	arrParam(5) = "�����ڵ�"				' �����ʵ��� �� ��Ī
		
	arrField(0) = "A.BANK_NM"					' Field��(1)
	arrField(1) = "B.BANK_ACCT_NO"				' Field��(2)
   
	arrHeader(0) = "�����"						' Header��(1)
	arrHeader(1) = "�������ڵ�"

	IsOpenPop = True
		        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBankAcct(arrRet)
	End If	
	
End Function

'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetBankAcct(byval arrRet)
	With frm1
		.vspdData2.Col = C_BankAcctCd
		.vspdData2.Text = arrRet(1)
	  
	    Call vspdData2_Change(.vspdData2.Col, .vspdData2.Row)				 ' ������ �о�ٰ� �˷��� 
	End With
End Function

'=======================================================================================================
'	Name : OpenBankAcct()
'	Description : Bank Account No PopUp
'=======================================================================================================
Function OpenPopup(Byval strCode ,Byval iWhere)
	Dim arrRet
	Dim IntRetCD
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then Exit Function

	Select Case iWhere
		Case 1
			'��Ƽ���� �ڻ��ڵ带 ������� ���
			If Trim(frm1.txtDeptCd.value) = ""  then
				IsOpenPop = False
				IntRetCD = DisplayMsgBox("127800",parent.VB_INFORMATION,"x","x")            '��: Display Message(There is no changed data.)
				frm1.txtDeptCd.focus
				Exit Function
			End If

			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = C_AsstNo

			arrParam(0) = "�ڻ��ڵ� �˾�"    ' �˾� ��Ī
			arrParam(1) = "a_asset_inform_of_dept a, a_asset_master b "    ' TABLE ��Ī
			arrParam(2) = strCode      ' Code Condition
			arrParam(3) = ""       ' Name Condition
			arrParam(4) = "	a.asst_no = b.asst_no and  a.dept_cd =  " & FilterVar(frm1.txtDeptCd.value, "''", "S") & "  and a.org_change_id =  " & FilterVar(frm1.hORGCHANGEID.value, "''", "S") & " "       ' Where Condition
			arrParam(5) = "�ڻ��ڵ�"     ' �����ʵ��� �� ��Ī

			arrField(0) = "a.asst_no"     ' Field��(0)
			arrField(1) = "b.asst_nm"     ' Field��(0)
			arrField(2) = "b.reg_dt"     ' Field��(0)
			arrField(3) = "acq_amt"     ' Field��(0)
			arrField(4) = "b.acq_loc_amt"     ' Field��(0)
			arrField(5) = "b.acq_qty"     ' Field��(0)
			arrField(6) = "a.inv_qty"     ' Field��(1)

			arrHeader(0) = "�ڻ��ڵ�"   ' Header��(0)
			arrHeader(1) = "�ڻ��"    ' Header��(1)
			arrHeader(2) = "�������"    ' Header��(1)
			arrHeader(3) = "���ݾ�"    ' Header��(1)
			arrHeader(4) = "���ݾ�(�ڱ�)"    ' Header��(1)
			arrHeader(5) = "������"    ' Header��(1)
			arrHeader(6) = "������"    ' Header��(1)
		Case 2
			arrParam(0) = "�Ա������˾�"    ' �˾� ��Ī
			arrParam(1) = " ( SELECT MINOR_CD, MINOR_NM FROM (SELECT A.MINOR_CD MINOR_CD, A.MINOR_NM MINOR_NM FROM B_MINOR A, B_CONFIGURATION B " & _
						  " WHERE (A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " ) " & _
						  " AND A.MINOR_CD NOT IN ( " & FilterVar("NP", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("AP", "''", "S") & " , " & FilterVar("CP", "''", "S") & "  , " & FilterVar("NE", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 UNION ALL " & _
						  " SELECT " & FilterVar("AR", "''", "S") & "  MINOR_CD, " & FilterVar("�̼���", "''", "S") & "  MINOR_NM ) A ) B"
			arrParam(2) = strCode      ' Code Condition
			arrParam(3) = ""       ' Name Condition
			arrParam(4) = "" 
			arrParam(5) = "�Ա�����"     ' �����ʵ��� �� ��Ī

			arrField(0) = "B.MINOR_CD"     ' Field��(0)
			arrField(1) = "B.MINOR_NM"     ' Field��(1)

			arrHeader(0) = "�Ա������ڵ�"   ' Header��(0)
			arrHeader(1) = "�Ա�������"    ' Header��(1)
		Case 7
			arrParam(0) = "�̼��� �˾�"    ' �˾� ��Ī
			arrParam(1) = "a_jnl_acct_assn a, a_acct b"    ' TABLE ��Ī
			arrParam(2) = strCode      ' Code Condition
			arrParam(3) = ""       ' Name Condition
			arrParam(4) = "A.trans_type = " & FilterVar("AS006", "''", "S") & "  and A.Acct_cd = B.Acct_cd and Jnl_cd = " & FilterVar("AR", "''", "S") & " "       ' Where Condition
			arrParam(5) = "�����ڵ�"     ' �����ʵ��� �� ��Ī

			arrField(0) = "a.ACCT_CD"     ' Field��(0)
			arrField(1) = "b.ACCT_NM"     ' Field��(1)

			arrHeader(0) = "�����ڵ�"   ' Header��(0)
			arrHeader(1) = "�����ڵ��"    ' Header��(1)
		Case 8
			arrParam(0) = "�������ڵ� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK A, B_BANK_ACCT B, F_DPST C"			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "												' Where Condition'			
			arrParam(4) = arrParam(4) & "AND B.BANK_ACCT_NO = C.BANK_ACCT_NO "	
			arrParam(4) = arrParam(4) & "AND A.BANK_CD = C.BANK_CD "
			arrParam(4) = arrParam(4) & "AND C.DPST_FG IN (" & FilterVar("SV", "''", "S") & " ," & FilterVar("ET", "''", "S") & " ) "
			arrParam(5) = "�������ڵ�"				' �����ʵ��� �� ��Ī 
					
   			arrField(0) = "A.BANK_CD"					' Field��(1)
			arrField(1) = "A.BANK_NM"					' Field��(1)
			arrField(2) = "B.BANK_ACCT_NO"				' Field��(2)
	
			arrHeader(0) = "�����ڵ�"						' Header��(1)
			arrHeader(1) = "�����"						' Header��(1)
			arrHeader(2) = "�������ڵ�"
	End Select
	

		        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCommOn(arrRet, iWhere)
	End If	
	
End Function

'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetCommOn(byval arrRet,byval iWhere)
	Dim Row
	With frm1
		Select Case iWhere
			Case 1
				'��Ƽ���� �ڻ��ڵ带 ������� ���
				ggoSpread.Source = frm1.vspdData
				Row = frm1.vspdData.ActiveRow
				.vspdData.row = Row
				.vspdData.Col = C_AsstNo
				.vspdData.value = arrRet(0)
				.vspdData.Col = C_AsstNm
				.vspdData.value = arrRet(1)
				.vspdData.Col = C_AcqDt
				.vspdData.text = UniConvDateAToB(arrRet(2), parent.gServerDateFormat,gDateFormat)
				.vspdData.Col = C_AcqLocAmt
				.vspdData.value = UNIConvNum(Trim(arrRet(4)),0)
				.vspdData.Col = C_InvQty
				.vspdData.value = UNIConvNum(Trim(arrRet(6)),0)
				lgBlnFlgChgValue = True
			Case 2
				ggoSpread.Source = frm1.vspdData2
				Row = frm1.vspdData2.ActiveRow
				.vspdData2.row = Row
				.vspdData2.Col = C_RcptTypeCd
				.vspdData2.value = arrRet(0)
				.vspdData2.Col = C_RcptTypeNm
				.vspdData2.value = arrRet(1)  
				lgBlnFlgChgValue = True
				Call vspdData2_Change(C_RcptTypeCd, frm1.vspdData2.ActiveRow)
			Case 7
				ggoSpread.Source = frm1.vspdData2
				Row = frm1.vspdData2.ActiveRow
				.vspdData2.row = Row
				.vspdData2.Col = C_ArAcctCd
				.vspdData2.value = arrRet(0)
				.vspdData2.Col = C_ArAcctNm
				.vspdData2.value = arrRet(1)  
				lgBlnFlgChgValue = True
			Case 8
				ggoSpread.Source = frm1.vspdData2
				Row = frm1.vspdData2.ActiveRow
				.vspdData2.row = Row
				.vspdData2.Col = C_BankCd
				.vspdData2.value = arrRet(0)
				.vspdData2.Col = C_BankNm
				.vspdData2.value = arrRet(1)  
				.vspdData2.Col = C_BankAcctCd
				.vspdData2.value = arrRet(2)
				lgBlnFlgChgValue = True
		End Select
	End With
End Function


Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
      	
	If IsOpenPop = True  Then Exit Function	
	if frm1.Rb_Duse.checked = True then Exit Function

	arrHeader(0) = "�ΰ�������"						' Header��(0)
	arrHeader(1) = "�ΰ�����"						' Header��(1)
	arrHeader(2) = "�ΰ���Rate"
    
	arrField(0) = "B_Minor.MINOR_CD"							' Field��(0)
	arrField(1) = "B_Minor.MINOR_NM"							' Field��(1)
    arrField(2) = "F2" & parent.gColSep & "b_configuration.REFERENCE"	
    
	arrParam(0) = "�ΰ�������"						' �˾� ��Ī
	arrParam(1) = "B_Minor,b_configuration"				' TABLE ��Ī
	arrParam(2) = Trim(frm1.txtVatType.value)			' Code Condition
			
	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd = B_Minor.Major_Cd"	 
	arrParam(5) = "�ΰ�������"						' TextBox ��Ī	

	IsOpenPop = True
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVatType.focus
		Exit Function
	Else
		Call SetVatType(arrRet)
	End If	
	
End Function

'=======================================================================================================
'	Name : Setvattype()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetVatType(byval arrRet)
	frm1.txtVatType.Value    = arrRet(0)		
	frm1.txtVatTypeNm.Value    = arrRet(1)		
	frm1.txtVatRate.text    = arrRet(2)		
	Call txtVatType_OnChange
	
	frm1.txtVatType.focus
	
	lgBlnFlgChgValue = True
End Function

'===========================================================================
' Function Name : OpenReportAreaCd
' Function Desc : OpenReportAreaCd Reference Popup
'===========================================================================
Function OpenReportAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "�Ű����� �˾�"	
	arrParam(1) = "B_TAX_BIZ_AREA"				
	arrParam(2) = Trim(frm1.txtReportAreaCd.value)
	arrParam(3) = "" 
	arrParam(4) = ""
	arrParam(5) = "�Ű�����"			
	
    arrField(0) = "TAX_BIZ_AREA_CD"	
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "�Ű������ڵ�"		
    arrHeader(1) = "�Ű������"		

	IsOpenPop = True
        
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtReportAreaCd.focus
		Exit Function
	Else
		Call SetReportArea(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetReportArea()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetReportArea(byval arrRet)
	frm1.txtReportAreaCd.Value		= arrRet(0)		
	frm1.txtReportAreaNm.Value		= arrRet(1)		
	
	frm1.txtReportAreaCd.focus
		
	lgBlnFlgChgValue = True
End Function



Function OpenVatType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
      	
	If IsOpenPop = True  Then Exit Function	
	if frm1.Rb_Duse.checked = True then Exit Function

	arrHeader(0) = "�ΰ�������"						' Header��(0)
	arrHeader(1) = "�ΰ�����"						' Header��(1)
	arrHeader(2) = "�ΰ���Rate"
    
	arrField(0) = "B_Minor.MINOR_CD"							' Field��(0)
	arrField(1) = "B_Minor.MINOR_NM"							' Field��(1)
    arrField(2) = "F2" & parent.gColSep & "b_configuration.REFERENCE"	
    
	arrParam(0) = "�ΰ�������"						' �˾� ��Ī
	arrParam(1) = "B_Minor,b_configuration"				' TABLE ��Ī
	arrParam(2) = Trim(frm1.txtVatType.value)			' Code Condition
			
	arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd =b_configuration.minor_cd and " & _
	              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd = B_Minor.Major_Cd"	 
	arrParam(5) = "�ΰ�������"						' TextBox ��Ī	

	IsOpenPop = True
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVatType.focus
		Exit Function
	Else
		Call SetVatType(arrRet)
	End If	
	
End Function

'=======================================================================================================
'	Name : Setvattype()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetVatType(byval arrRet)
	frm1.txtVatType.Value    = arrRet(0)		
	frm1.txtVatTypeNm.Value    = arrRet(1)		
	frm1.txtVatRate.text    = arrRet(2)		
	Call txtVatType_OnChange
	
	frm1.txtVatType.focus
	
	lgBlnFlgChgValue = True
End Function

'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================
Function OpenPopupTempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'������ǥ��ȣ 
	arrParam(1) = ""							'Reference��ȣ 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function
'=======================================================================================================
'Description : ȸ����ǥ �������� �˾�
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = ""						'Reference��ȣ 

	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================================================================================= 
'	Name : OpenAcctPopup()
'	Description : Ref ȭ���� call�Ѵ�. 
'========================================================================================================= 
Function OpenAsstPopup()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	If Trim(frm1.txtChgDt.text) = "" Then Exit Function
	
	iCalledAspName = AskPRAspName("A7127RA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A7127RA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If gSelframeFlg <> TAB1 Then Exit Function
	
	IsOpenPop = True

	'�μ��ڵ�� ����������̵� ������ ����.
	arrParam(0) = Trim(frm1.txtDeptCd.value)				' �˻������� ������� �Ķ���� 
	arrParam(1) = Trim(frm1.hORGCHANGEID.value)
	arrParam(2) = Trim(frm1.txtChgDt.text)
	
	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0,0) = "" Then	
		Exit Function
	Else
		Call SetRefOpenAsst(arrRet)
	End If
End Function

'========================================================================================================= 
'	Name : SetRefOpenAp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'========================================================================================================= 
Function SetRefOpenAsst(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim X
	Dim sFindFg

	With frm1
		.vspddata.focus		
		ggoSpread.Source = .vspddata
		.vspddata.ReDraw = False	
	
		TempRow = .vspddata.MaxRows												'��: ��������� MaxRows
		For I = TempRow to TempRow + Ubound(arrRet, 1)

			sFindFg	= "N"
			For x = 1 to TempRow
				.vspddata.Row = x
				.vspddata.Col = C_AsstNo
				If "" & UCase(Trim(.vspddata.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 0))) Then
					.vspddata.Col = C_DeptCd
					If "" & UCase(Trim(.vspddata.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 3))) Then
						.vspddata.Col = C_SubNo
						If "" & UCase(Trim(.vspddata.Text)) = "" & UCase(Trim(arrRet(I - TempRow, 2))) Then
							sFindFg	= "Y"
						End If
					End If
				End If
			Next

			If 	sFindFg	= "N" Then
				.vspddata.MaxRows = .vspddata.MaxRows + 1
				.vspddata.Row = I + 1				
				.vspddata.Col = 0

				.vspddata.Text = ggoSpread.InsertFlag
				.vspddata.Col = C_AsstNo '�ڻ��ȣ
				.vspddata.text = arrRet(I - TempRow, 0)
				.vspddata.Col = C_AsstNm '�ڻ��
				.vspddata.text = arrRet(I - TempRow, 1)
				.vspddata.Col = C_SubNo 'Sub No
				.vspddata.text = arrRet(I - TempRow, 2)
				.vspddata.Col = C_DeptCd '�μ��ڵ�
				.vspddata.text = arrRet(I - TempRow, 3)
				.vspddata.Col = C_DeptNm '�μ���
				.vspddata.text = arrRet(I - TempRow, 4)
				.vspddata.Col = C_OrgChgId ' ��������ID
				.vspddata.text = arrRet(I - TempRow, 14)
				.vspddata.Col = C_AcqDt '�������
				.vspddata.text = arrRet(I - TempRow, 7)
				.vspddata.Col = C_InvQty '������
				.vspddata.text = arrRet(I - TempRow, 12)
				.vspddata.Col = C_ChgQty '��������
				.vspddata.text = arrRet(I - TempRow, 12)
				.vspddata.Col = C_AcqLocAmt '����ѱݾ�
				.vspddata.text = arrRet(I - TempRow, 8)
				.vspddata.Col = C_DeprLocAmt '�����󰢴����
				.vspddata.text = arrRet(I - TempRow, 9)
				.vspddata.Col = C_BALLocAmt '��ΰ���
				.vspddata.text = arrRet(I - TempRow, 10)
				.vspddata.Col = C_SoldRate
				.vspddata.text = 100
				.vspddata.Lock = False
				ggoSpread.SpreadUnLock		C_SoldRate,		I + 1,		C_SoldRate
				ggoSpread.SSSetRequired		C_SoldRate,		I + 1,		C_SoldRate
				ggoSpread.SpreadUnLock		C_ChgQty,		I + 1,		C_ChgQty
				ggoSpread.SSSetRequired		C_ChgQty,		I + 1,		C_ChgQty
				ggoSpread.SpreadUnLock		C_ChgAmt,		I + 1,		C_ChgAmt
				ggoSpread.SSSetRequired		C_ChgAmt,		I + 1,		C_ChgAmt
				ggoSpread.SpreadUnLock		C_ChgLocAmt,		I + 1,		C_ChgLocAmt
				ggoSpread.SpreadUnLock		C_TaxAmt,		I + 1,		C_TaxAmt
				ggoSpread.SpreadUnLock		C_TaxLocAmt,		I + 1,		C_TaxLocAmt
				ggoSpread.SpreadUnLock		C_AsstSoldDesc,		I + 1,		C_AsstSoldDesc
				'ggoOper.SetReqAttr frm1.Rb_Sold,		 "Q"    '�ŷ�ó
				'ggoOper.SetReqAttr frm1.Rb_Duse,		 "Q"    '�ŷ�ó

			End If	
		Next	
		
		.vspddata.ReDraw = True
		If .vspddata.MaxRows	 > 0 Then
			Call ggoOper.SetReqAttr(frm1.txtChgDt, "Q")
			'Call ggoOper.SetReqAttr(frm1.txtChgDt, "N")
		End If

    End With
    
	if frm1.Rb_Duse.checked = true then
		frm1.txtRadio.value = "03"
		call Radio2_onChange()
	end if 
	
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

   ' ------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim strCodeList
	Dim strNameList
	ggoSpread.Source = frm1.vspdData2
	Call CommonQueryRs("A.MINOR_CD,A.MINOR_NM","B_MINOR A, B_CONFIGURATION B", _
					   "(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " ) AND A.MINOR_CD NOT IN ( " & FilterVar("NP", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("AP", "''", "S") & " , " & FilterVar("CP", "''", "S") & "  , " & FilterVar("NE", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 ", _
	                   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	'A1006

	strCodeList = Replace(lgF0, Chr(11), vbTab) & "AR"
	strNameList = Replace(lgF1, Chr(11), vbTab) & "����ä��"

    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	Dim indx
	Dim iRow
	Dim varData

	ggoSpread.Source = gActiveSpdSheet
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"		
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			Call SetSpreadColor(-1,-1,"A")

		Case "VSPDDATA2"			
			Call ggoSpread.RestoreSpreadInf()						
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()			
			Call InitData()

	End Select
End Sub


'======================================================================================================
'												3. Event��
'	���: Event �Լ��� ���� ó��
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ�
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     'Load table , B_numeric_format
        
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
                                                                            'Format Numeric Contents Field                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call InitSpreadSheet("A")                                                    'Setup the Spread sheet
    Call InitSpreadSheet("B")                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
    Call SetDefaultVal
	frm1.hORGCHANGEID.value =parent.gChangeOrgId 
    frm1.txtRadio.value = "03"
    
    Call SetToolBar("1110100100000111")										' ó�� �ε�� ǥ�� �� ���� 
   	lgBlnFlgChgValue = False

    frm1.txtAsstChgNo.focus 
	call txtDocCur_OnChangeASP()  
	
	' ���Ѱ��� �߰�
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' �����
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ�
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' ����
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing	

End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtChgDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgDt.Action = 7
    End If
End Sub

Sub txtIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDt.Action = 7
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPrpaymDt_Change()
'   Event Desc : 
'=======================================================================================================

Sub txtIssuedDt_Change()
    lgBlnFlgChgValue = True
End Sub



sub hORGCHANGEID_onchange()
	msgbox frm1.hORGCHANGEID.value 
end sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )

    With frm1.vspdData 
    End With
                
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	With frm1.vspdData
	
		If Col = C_ChgAmt then
			Call FncSumChgAmt()
		End If
		If Col = C_ChgLocAmt then
			Call FncSumChgLocAmt()
		End If
		If Col = C_TaxAmt then
			Call FncSumTaxAmt()
		End If
		If Col = C_TaxLocAmt then
			Call FncSumTaxLocAmt()
		End If

		If Col = C_ChgQty then 'jsk 2003/09/23
			.col = C_SoldRate
			.text = 0
		End If
		ggoSpread.Source = frm1.vspdData
		ggoSpread.UpdateRow Row   
		
	End With

End Sub

Sub FncSumChgAmt()
	Dim i
	Dim SumChgAmt
	
	SumChgAmt = 0
	
	With frm1.vspdData
		.Col = C_ChgAmt
		For i = 1 to frm1.vspdData.Maxrows
			.Row = i

			SumChgAmt = SumChgAmt + UNICDbl(.text)
		Next
	End With
	frm1.txtTotalAmt.text = UNIConvNumPCToCompanyByCurrency(SumChgAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	
End Sub


Sub FncSumChgLocAmt()
	Dim i
	Dim SumChgLocAmt
	
	SumChgLocAmt = 0
	
	With frm1.vspdData
		.Col = C_ChgLocAmt
		For i = 1 to frm1.vspdData.Maxrows
			.Row = i
			SumChgLocAmt = SumChgLocAmt + UNICDbl(.text)

		Next
	End With
	
	frm1.txtTotalLocAmt.text = UNIConvNumPCToCompanyByCurrency(SumChgLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	
End Sub


Sub FncSumTaxAmt()
	Dim i
	Dim SumTaxAmt
	
	SumTaxAmt = 0
	
	With frm1.vspdData
		.Col = C_TaxAmt
		For i = 1 to frm1.vspdData.Maxrows
			.Row = i
			SumTaxAmt = SumTaxAmt + UNICDbl(.text)

		Next
	End With
	frm1.txtVatAmt.text = UNIConvNumPCToCompanyByCurrency(SumTaxAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	
End Sub

Sub FncSumTaxLocAmt()
	Dim i
	Dim SumTaxLocAmt
	
	SumTaxLocAmt = 0
	
	With frm1.vspdData
		.Col = C_TaxLocAmt
		For i = 1 to frm1.vspdData.Maxrows
			.Row = i
			SumTaxLocAmt = SumTaxLocAmt + UNICDbl(.text)

		Next
	End With
	frm1.txtVatLocAmt.text = UNIConvNumPCToCompanyByCurrency(SumTaxLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	
End Sub



Sub FncSumRcptAmt()
	Dim i
	Dim SumRcptAmt
	
	SumRcptAmt = 0
	
	With frm1.vspdData2
		.Col = C_RcptAmt
		For i = 1 to frm1.vspdData2.Maxrows
			.Row = i
			SumRcptAmt = SumRcptAmt + UNICDbl(.text)

		Next
	End With
	frm1.txtTotalRcptAmt.text = UNIConvNumPCToCompanyByCurrency(SumRcptAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	
End Sub

Sub FncSumRcptLocAmt()
	Dim i
	Dim SumRcptLocAmt
	
	SumRcptLocAmt = 0
	
	With frm1.vspdData2
		.Col = C_RcptLocAmt
		For i = 1 to frm1.vspdData2.Maxrows
			.Row = i
			SumRcptLocAmt = SumRcptLocAmt + UNICDbl(.text)

		Next
	End With
	frm1.txtTotalRcptLocAmt.text = UNIConvNumPCToCompanyByCurrency(SumRcptLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )
	Dim intIndex
	Dim i
	
	On Error Resume Next
	Err.Clear                                                               '��: Clear error no

	lgBlnFlgChgValue = True
	ggoSpread.Source = frm1.vspdData2

    Select Case Col
		Case  C_RcptAmt			' ���뱸��
			Call FncSumRcptAmt()
		Case  C_RcptLocAmt			' ���뱸��
			Call FncSumRcptLocAmt()
		Case  C_RcptTypeCd			' ���뱸��
			
			frm1.vspdData2.ReDraw = False

			for i = C_ARAPNo to C_NoteNo
				frm1.vspdData2.col = i
				frm1.vspdData2.text = ""
			next 
			
			frm1.vspdData2.Col = C_RcptTypeCd
			
			Select Case frm1.vspdData2.text
			
			Case "AR"
				ggoSpread.SpreadUnLock		C_ARAPNo,		Row, Row
				ggoSpread.SpreadUnLock		C_ArAcctCd,		Row, Row
				ggoSpread.SSSetRequired		C_ArAcctCd,		Row, Row
				ggoSpread.SpreadUnLock		C_ArAcctPopup,	Row, Row
				ggoSpread.SSSetProtected	C_ArAcctNm,		Row, Row
				ggoSpread.SpreadUnLock		C_ArDueDt,		Row, Row
				ggoSpread.SSSetRequired		C_ArDueDt,		Row, Row
				frm1.vspdData2.Col = C_ArDueDt
				frm1.vspdData2.text = frm1.txtChgDt.text
				ggoSpread.SSSetProtected	C_BankCd,		Row, Row
				ggoSpread.SSSetProtected	C_BankPopup,	Row, Row
				ggoSpread.SSSetProtected	C_BankNm,		Row, Row
				ggoSpread.SSSetProtected	C_BankAcctCd,	Row, Row
				ggoSpread.SSSetProtected	C_NoteNo,		Row, Row
				ggoSpread.SSSetProtected	C_NotePopup,		Row, Row

			Case "DP"
				ggoSpread.SSSetProtected	C_ARAPNo,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctCd,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctPopup,	Row, Row
				ggoSpread.SSSetProtected	C_ArAcctNm,		Row, Row
				ggoSpread.SSSetProtected	C_ArDueDt,		Row, Row
				ggoSpread.SSSetRequired		C_BankCd,		Row, Row
				ggoSpread.SpreadUnLock		C_BankPopup,	Row, Row
				ggoSpread.SSSetProtected	C_BankNm,		Row, Row
				ggoSpread.SSSetProtected	C_BankAcctCd,	Row, Row
				ggoSpread.SSSetProtected	C_NoteNo,		Row, Row
				ggoSpread.SSSetProtected		C_NotePopup,		Row, Row

			Case "CS"
				ggoSpread.SSSetProtected	C_ARAPNo,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctCd,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctPopup,	Row, Row
				ggoSpread.SSSetProtected	C_ArAcctNm,		Row, Row
				ggoSpread.SSSetProtected	C_ArDueDt,		Row, Row
				ggoSpread.SSSetProtected	C_BankCd,		Row, Row
				ggoSpread.SSSetProtected	C_BankPopup,	Row, Row
				ggoSpread.SSSetProtected	C_BankNm,		Row, Row
				ggoSpread.SSSetProtected	C_BankAcctCd,	Row, Row
				ggoSpread.SSSetProtected	C_NoteNo,		Row, Row
				ggoSpread.SSSetProtected		C_NotePopup,		Row, Row

			Case "CK"
				ggoSpread.SSSetProtected	C_ARAPNo,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctCd,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctPopup,	Row, Row
				ggoSpread.SSSetProtected	C_ArAcctNm,		Row, Row
				ggoSpread.SSSetProtected	C_ArDueDt,		Row, Row
				ggoSpread.SSSetProtected	C_BankCd,		Row, Row
				ggoSpread.SSSetProtected	C_BankPopup,	Row, Row
				ggoSpread.SSSetProtected	C_BankNm,		Row, Row
				ggoSpread.SSSetProtected	C_BankAcctCd,	Row, Row
				ggoSpread.SSSetProtected	C_NoteNo,		Row, Row
				ggoSpread.SSSetProtected		C_NotePopup,		Row, Row

			Case else
				ggoSpread.SSSetProtected	C_ARAPNo,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctCd,		Row, Row
				ggoSpread.SSSetProtected	C_ArAcctPopup,	Row, Row
				ggoSpread.SSSetProtected	C_ArAcctNm,		Row, Row
				ggoSpread.SSSetProtected	C_ArDueDt,		Row, Row
				ggoSpread.SSSetProtected	C_BankCd,		Row, Row
				ggoSpread.SSSetProtected	C_BankPopup,	Row, Row
				ggoSpread.SSSetProtected	C_BankNm,		Row, Row
				ggoSpread.SSSetProtected	C_BankAcctCd,	Row, Row
				ggoSpread.SSSetRequired		C_NoteNo,		Row, Row
				ggoSpread.SpreadUnLock		C_NotePopup,		Row, Row
				
			End Select 
			
			frm1.vspdData2.ReDraw = False

	End Select 	
	
	ggoSpread.source = frm1.vspdData2
	
	frm1.vspdData2.row = Row
	frm1.vspdData2.col = 0
	
	If frm1.vspdData2.Text <> ggoSpread.DeleteFlag and frm1.vspdData2.Text <> ggoSpread.InsertFlag then
		frm1.vspdData2.Text = ggoSpread.UpdateFlag
	End If

End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻�
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
	gMouseClickStatus = "SPC"	'Split �����ڵ�
	   
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

	If Row <= 0 Then
	   ggoSpread.Source = frm1.vspdData
	   Exit Sub
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) �÷� width ���� �̺�Ʈ �ڵ鷯
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'======================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻�
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
	gMouseClickStatus = "SP2C"	'Split �����ڵ�
	   
    Set gActiveSpdSheet = frm1.vspdData2

	If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
		ggoSpread.Source = frm1.vspdData2
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

	If Row <= 0 Then
	   ggoSpread.Source = frm1.vspdData2
	   Exit Sub
	End If
End Sub
'========================================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) �÷� width ���� �̺�Ʈ �ڵ鷯
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(Col1,Col2)
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' �ŷ���ȭ�ϰ� Company ��ȭ�� �ٸ��� ȯ���� 0���� ����
		frm1.txtXchRate.text	= "0"                         ' ����Ʈ���� 1�� �� ������ ȯ���� �Էµ� ������ �Ǵ��Ͽ�
								                                        ' ȯ�������� ���� �ʰ� �Էµ� ������ ���. 
	Else 

		frm1.txtXchRate.text	= "1"
	End If	
	call txtDocCur_OnChangeASP()  
End Sub

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================
Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtChgDt.Text) = "" Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtChgDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			If lgIntFlgMode <> parent.OPMD_UMODE Then
				IntRetCD = DisplayMsgBox("124600","X","X","X")  
			End If
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))
				frm1.hORGCHANGEID.value = Trim(arrVal2(2))
			Next	
			
		End If

End Sub

'==========================================================================================
'   Event Name : DeptCd_underChange(Byval strCode)
'   Event Desc : 
'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 

    If Trim(frm1.txtChgDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S")
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtChgDt.Text, parent.gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.value = ""
		frm1.txtDeptNm.value = ""
		frm1.hORGCHANGEID.value	= ""
	
	End If 
	'----------------------------------------------------------------------------------------

End Sub


Sub txtChgDt_onBlur()
    
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtChgDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hORGCHANGEID.value) Then
					'IntRetCD = DisplayMsgBox("124600","X","X","X") 
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hORGCHANGEID.value = ""
					.txtDeptCd.focus
			End if
		End If
	End With
'----------------------------------------------------------------------------------------

End Sub
'==========================================================================================
'   Event Name : txtVatType_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtVatType_OnChange()
	Dim dblVatAmt
	
	lgBlnFlgChgValue = True
	
	if frm1.txtVatAmt.text = "" then
		dblVatAmt = 0
	else
		dblVatAmt = UNICDbl(frm1.txtVatAmt.text)	
	end if
	
	If Trim(frm1.txtVatType.Value) = "" and dblVatAmt = 0 Then
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '�ΰ���Ÿ��
	Else
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '�ΰ���Ÿ��
	End If

End Sub


'==========================================================================================
'   Event Name : txtVatAmt_Change
'   Event Desc : 
'==========================================================================================
Sub txtVatAmt_Change()
	Dim dblVatAmt

	lgBlnFlgChgValue = True	
	
	if frm1.txtVatAmt.text="" then
		dblVatAmt = 0
	else
		dblVatAmt = UNICDbl(frm1.txtVatAmt.text)	
	end if
		
	If dblVatAmt = 0 and Trim(frm1.txtVatType.Value) = "" Then
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '�ΰ���Ÿ��
	Else
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '�ΰ���Ÿ�� 
	End IF
		
End Sub

'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChangeASP()
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	END IF	    
End Sub

'==========================================================================================
'   Event Name : txtCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtCur_OnChange()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ�
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
'   Event Desc : Spread Split �����ڵ�
'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

Sub subVspdSettingChange(ByVal lRow, Byval varData)	
	ggoSpread.Source = frm1.vspdData2
		
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)				
				Case "DP" & Chr(11)			' ������
					ggoSpread.SSSetRequired	 C_BankAcctCd,		 lRow, lRow			
					ggoSpread.SpreadUnLock   C_BankAcctcd,      lRow, C_BankAcctcd
					ggoSpread.SpreadUnLock   C_ArAcctPopup, lRow, C_ArAcctPopup
					ggoSpread.SSSetEdit		 C_BankAcctcd, "�������ڵ�", 25, 0, lRow, 30,2  
					ggoSpread.SSSetRequired	 C_BankAcctcd,      lRow, lRow	
					ggoSpread.SpreadLock     C_NoteNo,		 lRow, C_NoteNo,lRow   '������ȣ protect
					ggoSpread.SSSetProtected C_NoteNo,       lRow, lRow						
				Case "NO" & Chr(11)				
					ggoSpread.SpreadUnLock   C_NoteNo,        lRow, C_NoteNo,       lRow
					ggoSpread.SpreadLock     C_BankAcctcd,      lRow, C_BankAcctcd,     lRow   
					ggoSpread.SpreadLock     C_ArAcctPopup, lRow, C_ArAcctPopup,lRow
					ggoSpread.SSSetProtected C_BankAcctcd,      lRow, lRow								
					ggoSpread.SSSetEdit      C_NoteNo, "������ȣ", 25, 0, lRow, 30,2
					ggoSpread.SSSetRequired  C_NoteNo,        lRow, lRow
				Case Else									
					ggoSpread.SpreadLock     C_BankAcctcd,      lRow, C_BankAcctcd,     lRow   			
					ggoSpread.SpreadLock     C_ArAcctPopup, lRow, C_ArAcctPopup,lRow
					ggoSpread.SSSetProtected C_BankAcctcd,      lRow, lRow							
					ggoSpread.SpreadLock     C_NoteNo,        lRow, C_NoteNo,     lRow
					ggoSpread.SSSetProtected C_NoteNo,        lRow, lRow													
			End Select			
		
	End if
	
End Sub	

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻�
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻�
'=======================================================================================================
Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strData
	Dim strCard
    Dim strCode
	Dim strTemp
	With frm1.vspdData2 
	
		ggoSpread.Source = frm1.vspdData2
		Select Case Col
		Case C_RcptTypePopup 

			frm1.vspdData2.Col = C_RcptTypeCd
			frm1.vspdData2.Row = Row
			strCode = frm1.vspdData2.text
			Call OpenPopup(strCode ,2)		'�Ա�����
		
		Case C_ArAcctPopup 

			frm1.vspdData2.Col = C_ArAcctCd
			frm1.vspdData2.Row = Row
			strCode = frm1.vspdData2.text
			Call OpenPopup(strCode ,7)

		Case  C_BankPopup			' ���뱸��
			frm1.vspdData2.Col = C_BankCd
			frm1.vspdData2.Row = Row
			strCode = frm1.vspdData2.text
			Call OpenPopup(strCode ,8)
		Case  C_NotePopup			' ������ȣ
			frm1.vspdData2.Col = C_NoteNo
			frm1.vspdData2.Row = Row
			strTemp = Trim(.text)				    
			frm1.vspdData2.Col = C_RcptTypeCd
			strCard = frm1.vspdData2.text
			Call OpenNoteNo(strData, strCard)
		End Select
	
	End With
End Sub


Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With

End Sub

Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	 '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ����
			DbQuery
		End If
    End if
        
End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================


Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	 '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData2.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ
		If lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ����
			DbQuery
		End If
    End if
        
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
  '-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	    Exit Function
    	End If
    End If
  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    Call InitVariables                                                      'Initializes local global variables

	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery															'Query db data
       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                                          
	
	'-----------------------
	'Check previous data area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
    
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
	Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
	Call InitVariables                                                      'Initializes local global variables

	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
	Call SetDefaultVal
    call txtDocCur_OnChangeASP()  

    gSelframeFlg = TAB2
	Call ClickTab1()
    Call SetToolBar("1110100100100111")										' ó�� �ε�� ǥ�� �� ���� 

	lgBlnFlgChgValue = False	
	
	frm1.txtRadio.value = "03"
	
	if frm1.Rb_Duse.checked = True then    '�Ű��� ��,
		call Radio2_onChange()
	end if

	frm1.hORGCHANGEID.value =parent.gChangeOrgId 

	FncNew = True 

End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================

Function FncDelete() 
    Dim IntRetCD
	FncDelete = False
		
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")   '�����Ͻðڽ��ϱ�?  
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Precheck area
	'-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")                                
    	Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete                                                          '��: Delete db data
    
    FncDelete = True

End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	Dim lDelRows, intRows
	Dim iDx
	Dim lgvspdData
	Dim lgvspdData2
	
	FncSave = False

	ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer   
	lgvspdData = ggoSpread.SSCheckChange
	
	ggoSpread.Source = frm1.vspdData2                         '��: Preset spreadsheet pointer   
	lgvspdData2 = ggoSpread.SSCheckChange
	
    If lgBlnFlgChgValue = False and lgvspdData = False and lgvspdData2 = False  Then  '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
   
    If Not chkField(Document, "2") Then               '��: Check required field(Single area)
       Exit Function
    End If

	if frm1.vspdData.MaxRows < 1 then  
		IntRetCD = DisplayMsgBox("117370","X","X","X")  '�ڻ� �Ű���� ������ �ڻ��� ����Ͻʽÿ�.
		Exit Function
	end if

	if frm1.vspdData2.MaxRows < 1 and frm1.Rb_Sold.checked = true then   
		IntRetCD = DisplayMsgBox("117992","X","X","X")  '%1 �Ա� ������ �Է��Ͻʽÿ�.
		Exit Function
	end if

	Call FncSumChgAmt()
	Call FncSumChgLocAmt()
	Call FncSumTaxAmt()
	Call FncSumTaxLocAmt()
	Call FncSumRcptLocAmt()
	Call FncSumRcptLocAmt()
	
	'==================================================
	'	���뺯�ݾ� üũ
	If (UNICDbl(frm1.txtTotalAmt.text) <> UNICDbl(frm1.txtTotalRcptAmt.text)) Then
        IntRetCD = DisplayMsgBox("117380","x","x","x")            '%1 ���Ǹžװ� ���Աݾ��� ��ġ���� �ʽ��ϴ�..
       Exit Function
	End If
	'==================================================

	ggoSpread.Source = frm1.vspdData 
	For iDx = 1 To frm1.vspdData.MaxRows                        ' ����� üũ
		frm1.vspdData.Row = iDx
		frm1.vspdData.Col = C_AcqDt
		If UniConvDate(frm1.txtChgDt.text) < UniConvDate(frm1.vspdData.text) Then
			 IntRetCD = DisplayMsgBox("972002","x","�����","�Ű�/�����")		'%1 ��(��) %2 ���� ũ�ų� ���ƾ��մϴ�.
			Exit Function
		End IF

		'jsk 2003/09/23 C_SoldRate =0 �ΰ� ��ũ
		frm1.vspdData.Col = C_SoldRate
		If UNICDbl(frm1.vspdData.text) = 0 Then
			IntRetCD = DisplayMsgBox("141704","x","�Ű�����(%)","")		'%1 ��(��) 0 �� ���� �����ϴ�.
			Exit Function
		End If

	Next

	if frm1.Rb_Sold.checked = true then   
		ggoSpread.Source = frm1.vspdData2
		For iDx = 1 To frm1.vspdData2.MaxRows                         '�������� üũ
			frm1.vspdData2.Row = iDx
			frm1.vspdData2.Col = C_RcptTypeCd
			if frm1.vspdData2.text = "AR" Then
				frm1.vspdData2.Col = C_ArDueDt
				If UniConvDate(frm1.txtChgDt.text) >= UniConvDate(frm1.vspdData2.text) Then
					IntRetCD = DisplayMsgBox("972002","x","�̼��ݸ�����","�Ű������")
					Exit Function
				End IF
			End If 
		Next
	End If

	if frm1.Rb_Sold.checked = True then    '�Ű��� ��,
		ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
		If Not ggoSpread.SSDefaultCheck  Then              '��: Check required field(Multi area)
		   Exit Function
		End If
	else  ''''' ����� ��, grid�� �ڻ�����󼼳��� �Է½�,����
		if frm1.vspdData2.MaxRows > 0 then 
			ggoSpread.Source = frm1.vspdData2
			for intRow = 1 to frm1.vspdData2.MaxRows 				
				frm1.vspdData2.row = intRow
				lDelRows = ggoSpread.DeleteRow				
			next
			ggoSpread.Source = frm1.vspdData2
			ggospread.ClearSpreadData		'Buffer Clear
		end if
	end if

	'-----------------------
	'Save function call area
	'-----------------------
	Call DbSave				                                                '��: Save db data	
	FncSave = True
	
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()

	If gSelframeFlg = TAB1 Then
	
	    frm1.vspdData.ReDraw = False
		if frm1.vspdData.MaxRows < 1 then Exit Function
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow,"A"
	    frm1.vspdData.ReDraw = True
	
	Else
    
		frm1.vspdData2.ReDraw = False
		if frm1.vspdData2.MaxRows < 1 then Exit Function
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.CopyRow
		SetSpreadColor frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow,"B"
		frm1.vspdData2.ReDraw = True
	
	End If    
	    	
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
	
	Dim iDx
	
	FncCancel = False

	If gSelframeFlg = TAB1 Then
	    if frm1.vspdData.MaxRows < 1 then Exit Function
		ggoSpread.Source = frm1.vspdData
		ggoSpread.EditUndo
	Else
	    if frm1.vspdData2.MaxRows < 1 then Exit Function
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.EditUndo
	End If    

    Set gActiveElement = document.ActiveElement   
     
    FncCancel = True
	
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(Byval pvRowCnt)

	Dim imRow, indx
	
	FncInsertRow = False
	
	If gSelframeFlg = TAB1 Then
		if IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		else
			imRow = AskSpdSheetAddRowcount()
			If ImRow="" then
				Exit Function
			End If
		End If

		With frm1
			.vspdData.focus
			ggoSpread.Source = .vspdData
			.vspdData.ReDraw = False
			ggoSpread.InsertRow ,imRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1,"A"
			.vspdData.ReDraw = True	
		End With
	Else
		if IsNumeric(Trim(pvRowCnt)) Then
			imRow = CInt(pvRowCnt)
		else
			imRow = AskSpdSheetAddRowcount()
			If ImRow="" then
				Exit Function
			End If
		End If

		With frm1
			if frm1.Rb_Sold.checked = True then		
				.vspdData2.focus
				ggoSpread.Source = .vspdData2
				.vspdData2.ReDraw = False
				ggoSpread.InsertRow ,imRow
				SetSpreadColor .vspdData2.ActiveRow, .vspdData2.ActiveRow + imRow - 1,"B"
				.vspdData2.ReDraw = True	
   			end if	
		End With
	End If    
    Set gActiveElement = document.ActiveElement  
	
	If Err.number = 0 Then
	   FncInsertRow = True                                                          '��: Processing is OK
	End If 
	
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
	
	if gSelframeFlg = TAB1 then 
	    if frm1.vspdData.MaxRows < 1 then Exit Function

	    With frm1.vspdData 
	    	.focus
	    	ggoSpread.Source = frm1.vspdData 
	    	lDelRows = ggoSpread.DeleteRow
	    End With
	else 
	    if frm1.vspdData2.MaxRows < 1 then Exit Function

	    With frm1.vspdData2 
	    	.focus
	    	ggoSpread.Source = frm1.vspdData2
	    	lDelRows = ggoSpread.DeleteRow
	    End With
	end if
	
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)										
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab����
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
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


'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete() 
    Dim strVal
    
    DbDelete = False														'��: Processing is NG 
    
     Call LayerShowHide(1)  
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtAsstChgNo2=" & Trim(frm1.txtAsstChgNo2.value)			'��: ���� ���� ����Ÿ
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����
    
    DbDelete = True                                                         '��: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ����
'========================================================================================================
Function DbDeleteOk()												        '���� ������ ���� ����
	lgBlnFlgChgValue = False
	Call FncNew()
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
    
	DbQuery = False
	
	Call LayerShowHide(1)
	
	Dim strVal
	With frm1
        	strVal = BIZ_LOAD_ID & "?txtMode=" & parent.UID_M0001						'��: 
        	strVal = strVal     & "&txtAsstChgNo=" & Trim(.txtAsstChgNo.value)	'��ȸ ���� ����Ÿ
        	strVal = strVal     & "&lgtab=" & gSelframeFlg
        	strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey
        	strVal = strVal     & "&lgStrPrevKey2=" & lgStrPrevKey2
	End With

	' ���Ѱ��� �߰�
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' �����
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ�
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����        	

	Call RunMyBizASP(MyBizASP, strVal)										'�����Ͻ� ASP �� ����
	
	DbQuery = True
    
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű�
'========================================================================================================
Function DbQueryOk()													'��ȸ ������ �������
	Dim varData
	Dim iRow
	
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
	Call SetToolBar("1111111100111111")									'��ư ���� ����
	
	Call InitData()


	With frm1
		.vspdData.Redraw = False
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SSSetProtected	C_ChgQty,		-1, C_ChgQty
		ggoSpread.SSSetProtected	C_SoldRate,		-1, C_SoldRate
		ggoSpread.SSSetProtected	C_ChgAmt,		-1, C_ChgAmt
		ggoSpread.SSSetProtected	C_ChgLocAmt,		-1, C_ChgLocAmt
		ggoSpread.SSSetProtected	C_TaxAmt,		-1, C_TaxAmt
		ggoSpread.SSSetProtected	C_TaxLocAmt,		-1, C_TaxLocAmt
		ggoSpread.SSSetProtected	C_TaxLocAmt,		-1, C_TaxLocAmt
		ggoSpread.SSSetProtected	C_AsstSoldDesc,		-1, C_AsstSoldDesc
		.vspdData.Redraw = True
	If frm1.vspdData.MaxRows > 0 Then
		Call ggoOper.SetReqAttr(frm1.txtChgDt, "Q")
	End If

		.vspdData2.Redraw = False
		ggoSpread.Source = frm1.vspdData2
		For iRow = 1 To frm1.vspdData2.MaxRows
	
			.vspdData2.Col = C_RcptTypeCd
			.vspdData2.Row = iRow
				
			select case frm1.vspdData2.text

			Case "AR"

				ggoSpread.SSSetProtected	C_ARAPNo,		iRow, iRow

				ggoSpread.SSSetRequired		C_ArAcctCd,		iRow, iRow
				ggoSpread.SpreadUnLock		C_ArAcctPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctNm,		iRow, iRow
				ggoSpread.SSSetRequired		C_ArDueDt,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_BankNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankAcctCd,	iRow, iRow

				ggoSpread.SSSetProtected	C_NoteNo,		iRow, iRow
				ggoSpread.SSSetProtected	C_NotePopup,	iRow, iRow

			Case "DP"
				ggoSpread.SSSetProtected	C_ARAPNo,		iRow, iRow

				ggoSpread.SSSetProtected	C_ArAcctCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArDueDt,		iRow, iRow

				ggoSpread.SSSetRequired		C_BankCd,		iRow, iRow
				ggoSpread.SpreadUnLock		C_BankPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_BankNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankAcctCd,	iRow, iRow
				ggoSpread.SSSetProtected	C_NoteNo,		iRow, iRow
				ggoSpread.SSSetProtected	C_NotePopup,	iRow, iRow

			Case "CS"
				ggoSpread.SSSetProtected	C_ARAPNo,		iRow, iRow

				ggoSpread.SSSetProtected	C_ArAcctCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArDueDt,		iRow, iRow

				ggoSpread.SSSetProtected	C_BankCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_BankNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankAcctCd,	iRow, iRow

				ggoSpread.SSSetProtected	C_NoteNo,		iRow, iRow
				ggoSpread.SSSetProtected	C_NotePopup,	iRow, iRow

			Case "CK"
				ggoSpread.SSSetProtected	C_ARAPNo,		iRow, iRow

				ggoSpread.SSSetProtected	C_ArAcctCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArDueDt,		iRow, iRow

				ggoSpread.SSSetProtected	C_BankCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_BankNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankAcctCd,	iRow, iRow

				ggoSpread.SSSetProtected	C_NoteNo,		iRow, iRow
				ggoSpread.SSSetProtected	C_NotePopup,	iRow, iRow

			Case else
				ggoSpread.SSSetProtected	C_ARAPNo,		iRow, iRow

				ggoSpread.SSSetProtected	C_ArAcctCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_ArAcctNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_ArDueDt,		iRow, iRow

				ggoSpread.SSSetProtected	C_BankCd,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankPopup,	iRow, iRow
				ggoSpread.SSSetProtected	C_BankNm,		iRow, iRow
				ggoSpread.SSSetProtected	C_BankAcctCd,	iRow, iRow

				ggoSpread.SSSetRequired		C_NoteNo,		iRow, iRow
				ggoSpread.SpreadUnLock		C_NotePopup,	iRow, iRow
				
			End Select 

		Next
		
		.vspdData2.Redraw = True
	End With
	'call txtDocCur_OnChangeASP()
	'Call txtVatAmt_Change()
	'call txtVatType_OnChange()
	
	IF frm1.Rb_Duse.checked	= True Then
		frm1.txtRadio.value = "03"
		Call radio2_onchange()
	END IF

    'Call SetDefaultVal
	
	gSelframeFlg = TAB2
	Call ClickTab1()
	
	lgBlnFlgChgValue = False
	
End Function

Sub InitData()
End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave()
	
	Dim IntRows
	Dim IntCols
	
	Dim lGrpcnt
	Dim strVal
	Dim strDel
	
	Dim strAsstNo
	
	DbSave = False
	
	Call LayerShowHide(1)
	
	strVal = ""
	strDel = ""
	
	With frm1
		.txtMode.value = parent.UID_M0002									'��: ���� ����
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ����
	End With
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data ���� ��Ģ
	' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ
	
	lGrpCnt = 1

	With frm1.vspdData
	    
		For IntRows = 1 To .MaxRows

			.Row = IntRows

			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then

				.Col = C_ChgNo
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_AsstNo
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_DeptCd
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_OrgChgId
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_ChgQty
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_ChgAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_ChgLocAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_TaxAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_TaxLocAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_AsstSoldDesc
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_AcqLocAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_DeprLocAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_BALLocAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_SoldRate
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
				.Col = C_SubNo
				strVal = strVal & Trim(.Text) & parent.gRowSep
			
				lGrpCnt = lGrpCnt + 1
				
			End if
		Next

	End With

	frm1.txtMaxRows.value = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹��
	frm1.txtSpread.value = strVal									'��: Spread Sheet ������ ����

	strVal = "" 
	lGrpCnt = 1

	With frm1.vspdData2
	    
		For IntRows = 1 To .MaxRows
		
			.Row = IntRows
			
			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then

				.Col = C_RcptTypeCd
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_RcptAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep					
				.Col = C_RcptLocAmt
				strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep					
				.Col = C_ArAcctCd
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_ArDueDt
				strVal = strVal & UniConvDateToYYYYMMDD(.Text, parent.gDateFormat,"") & parent.gColSep
				.Col = C_ARAPNo
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_BankCd
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_BankAcctCd
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_NoteNo
				strVal = strVal & Trim(.Text) & parent.gColSep
				.Col = C_RcptDesc
				strVal = strVal & Trim(.Text) & parent.gRowSep					
							
				lGrpCnt = lGrpCnt + 1
			
			End If

		Next
	End With	
	
	With frm1
		.txtMaxRows2.value = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹��
		.txtSpread2.value = strVal									'��: Spread Sheet ������ ����
	
		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end	
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'��: ���� �����Ͻ� ASP �� ����

	DbSave = True                                                           
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű�
'=======================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ����

   	lgBlnFlgChgValue = false	

    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field

    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
	ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData		'Buffer Clear

    Call InitVariables                                                      'Initializes local global variables
	
	Call DbQuery	

End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Function Radio1_onChange()	

on error resume next
err.clear

	If frm1.txtRadio.value = "03" Then Exit function

	frm1.txtRadio.value = "03"
	
	ggoOper.SetReqAttr frm1.txtDocCur,	    "N"    '�ŷ���ȭ
	frm1.txtDocCur.value = parent.gCurrency
	ggoOper.SetReqAttr frm1.txtBpCd,		"N"    '�ŷ�ó

	ggoOper.SetReqAttr frm1.txtVatType,		"D"    '�ΰ�������
	ggoOper.SetReqAttr frm1.txtVatRate,		"D"    '�ΰ�����

	ggoOper.SetReqAttr frm1.txtReportAreaCd,"D"    '�Ű�����
	ggoOper.SetReqAttr frm1.txtIssuedDt,	"D"    '������
		

	frm1.txtIssuedDt.text	= UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,gDateFormat)	

	.vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLock	C_ChgNo,			-1,		C_ChgNo
	ggoSpread.SpreadLock	C_AsstNo,			-1,		C_AsstNo
	ggoSpread.SpreadLock	C_AsstNm,			-1,		C_AsstNm
	ggoSpread.SpreadLock	C_DeptCd,			-1,		C_DeptCd
	ggoSpread.SpreadLock	C_DeptNm,			-1,		C_DeptNm
	ggoSpread.SpreadLock	C_AcqDt,			-1,		C_AcqDt
	ggoSpread.SpreadLock	C_InvQty,			-1,		C_InvQty
	ggoSpread.SpreadUnLock	C_ChgQty,			-1,		C_ChgQty
	ggoSpread.SSSetRequired	C_ChgQty,			-1,		C_ChgQty	
	'ggoSpread.SpreadUnLock	C_SoldRate,			-1,		C_SoldRate
	'ggoSpread.SSSetRequired	C_SoldRate,			-1,		C_SoldRate
	ggoSpread.SpreadUnLock	C_ChgAmt,			-1,		C_ChgAmt
	ggoSpread.SSSetRequired	C_ChgAmt,			-1,		C_ChgAmt
	ggoSpread.SpreadUnLock	C_ChgLocAmt,		-1,		C_ChgLocAmt
	ggoSpread.SpreadLock	C_AcqLocAmt,		-1,		C_AcqLocAmt
	ggoSpread.SpreadLock	C_DeprLocAmt,		-1,		C_DeprLocAmt
	ggoSpread.SpreadLock	C_BALLocAmt,		-1,		C_BALLocAmt
	ggoSpread.SpreadLock	C_MnthDeprAmt,		-1,		C_MnthDeprAmt
	ggoSpread.SpreadUnLock	C_ChgLocAmt,		-1,		C_ChgLocAmt
	ggoSpread.SpreadLock	C_DeprLocAmt,		-1,		C_DeprLocAmt
	ggoSpread.SpreadUnLock	C_TaxAmt,			-1,		C_TaxAmt
	ggoSpread.SpreadUnLock	C_TaxLocAmt,		-1,		C_TaxLocAmt
	ggoSpread.SpreadLock	C_AccDeprAmt,		-1,		C_AccDeprAmt
			
	.vspdData.ReDraw = True

    Call InitSpreadSheet("B")                                                    'Setup the Spread sheet

    If lgIntFlgMode <> parent.OPMD_CMODE then                              'Indicates that current mode is Create mode
		Call SetToolBar("11111011100111111")									'��ư ���� ����	
		lgBlnFlgChgValue = True	
	Else
	    Call SetToolBar("1110100100111111")	
	End if

End Function

Function Radio2_onChange()

	Dim lDelRows,intRow, intCol
	Dim bMidChgVal

	Call ClickTab1()

	If frm1.txtRadio.value = "04" Then Exit function
	
	frm1.txtRadio.value = "04"
	
	ggoOper.SetReqAttr frm1.txtDocCur,		 "Q"    '�ŷ���ȭ
	ggoOper.SetReqAttr frm1.txtBpCd,		 "Q"    '�ŷ�ó

	ggoOper.SetReqAttr frm1.txtVatType,		 "Q"    '�ΰ��� ����
	ggoOper.SetReqAttr frm1.txtVatRate,		"Q"    '�ΰ�����

	ggoOper.SetReqAttr frm1.txtReportAreaCd, "Q"    '�Ű�����
	ggoOper.SetReqAttr frm1.txtIssuedDt,	 "Q"    '������

	bMidChgVal = lgBlnFlgChgValue

	frm1.txtBpCd.value = ""
	frm1.txtBpNm.value = ""
	frm1.txtDocCur.value = ""

	frm1.txtVatType.value = ""
	frm1.txtVatTypeNm.value = ""
	
	frm1.txtReportAreaCd.value = ""
	frm1.txtReportAreaNm.value = ""
	frm1.txtIssuedDt.text = ""

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = false	

	for intRow = 1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = intRow
			frm1.vspdData.Col = C_ChgAmt
			frm1.vspdData.text = ""
			ggoSpread.SSSetProtected	C_ChgAmt,		intRow, intRow	
			frm1.vspdData.Col = C_ChgLocAmt
			frm1.vspdData.text = ""
			ggoSpread.SSSetProtected	C_ChgLocAmt,		intRow, intRow	
			frm1.vspdData.Col = C_TaxAmt
			frm1.vspdData.text = ""
			ggoSpread.SSSetProtected	C_TaxAmt,		intRow, intRow	
			frm1.vspdData.Col = C_TaxLocAmt
			frm1.vspdData.text = ""
			ggoSpread.SSSetProtected	C_TaxLocAmt,		intRow, intRow	
	next
	frm1.vspdData.ReDraw = True

	if frm1.vspdData2.MaxRows > 0 then 
		ggoSpread.Source = frm1.vspdData2
		frm1.vspdData2.ReDraw = false	
		for intRow = 1 to frm1.vspdData2.MaxRows
			frm1.vspdData2.row = intRow
			lDelRows = ggoSpread.DeleteRow				
		next
		ggoSpread.Source = frm1.vspdData2
		ggospread.ClearSpreadData		'Buffer Clear
		
		frm1.vspdData.ReDraw = True
	end if

	frm1.txtTotalAmt.text = 0
	frm1.txtTotalRcptAmt.text = 0
	frm1.txtTotalLocAmt.text = 0
	frm1.txtTotalRcptLocAmt = 0
	frm1.txtVatAmt.text = 0
	frm1.txtVatLocAmt.text = 0
	frm1.txtDocCur.value = parent.gCurrency
	lgBlnFlgChgValue = bMidChgVal
	
    If lgIntFlgMode <> parent.OPMD_CMODE then                              'Indicates that current mode is Create mode
		Call SetToolBar("1111101100111111")									'��ư ���� ����	
		lgBlnFlgChgValue = True	
	Else
	    Call SetToolBar("1110100100111111")	
	End if
End Function

function txtDeptCd_onblur()
	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
end function

function txtBpCd_onblur()
	if Trim(frm1.txtBpCd.value) = "" then 		
		frm1.txtBpNm.value = ""		
	end if	
End function

Function txtDueDt_Change()
	lgBlnFlgChgValue = True
End Function


Function txtIssuedDt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtXchRate_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtTotalAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtTotalLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtVatRate_Change()
	lgBlnFlgChgValue = True
End Function

Function txtVatLocAmt_Change()
	lgBlnFlgChgValue = True
End Function

Function txtChgQty_Change()
	lgBlnFlgChgValue = True
End Function

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1

		ggoOper.FormatFieldByObjectOfCur .txtTotalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtTotalRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec

	End With

End Sub
'===================================== CurFormatNumericOCXRef()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCXRef()
End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'�Ǹž�
		ggoSpread.SSSetFloatByCellOfCur C_ChgAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec	
		ggoSpread.SSSetFloatByCellOfCur C_TaxAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec	
		

		ggoSpread.Source = frm1.vspdData2
		'�ݾ�
		ggoSpread.SSSetFloatByCellOfCur C_RcptAmt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--
'======================================================================================================
'       					6. Tag��
'	���: Tag�κ� ����
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="NO">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�Աݳ���</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><a href="vbscript:OpenAsstPopup()">�ڻ�����</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>					
			        <TD WIDTH=10>&nbsp;</TD>								
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
									<TD CLASS="TD5" NOWRAP>�Ű�����ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtAsstChgNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="�Ű�����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenChgNoInfo"></TD>
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
						        <TD CLASS="TD5" NOWRAP>�Ű�����ȣ</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAsstChgNo2" SIZE=20 MAXLENGTH=18 tag="25XXXU" ALT="�Ű�����ȣ"></TD>										        							
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
						        <TD CLASS="TD5" NOWRAP>����</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_Sold Checked tag = 23 value="03" onclick=radio1_onchange()><LABEL FOR=Rb_Sold>�Ű�</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Duse tag = 23 value="04" onclick=radio2_onchange()><LABEL FOR=Rb_Duse>���</LABEL></TD>										        							
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="����" tag="22X1" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDeptCd(frm1.txtDeptCd.Value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 tag="24" ALT="ȸ��μ���"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>								
<%	If gIsShowLocal <> "N" Then	%>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" TYPE="Text" SIZE=10 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurrency()"></TD>
								<TD CLASS="TD5" NOWRAP>ȯ��</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" id=fpDoubleSingle5></OBJECT>');</SCRIPT></TD>
							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur"><INPUT TYPE=HIDDEN NAME="txtXchRate">
<%	End If %>
							
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:call OpenBp(frm1.txtBpCd.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24" ALT="�ŷ�ó��"></TD>																		
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtChgDesc" SIZE=35 MAXLENGTH=30 tag="2X" ALT="����"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="������ǥ��ȣ"></TD>
								<TD CLASS="TD5" NOWRAP>ȸ����ǥ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=18 tag="24XXXU" ALT="ȸ����ǥ��ȣ"></TD>
							</TR>
							<TR>
								<TD WIDTH=100% HEIGHT= 100% VALIGN=TOP COLSPAN = 4>
								<DIV ID="TabDiv" STYLE="FlOAT: left; HEIGHT:100%; OVERFLOW:auto; WIDTH:100%;" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%> HEIGHT = 100%>
<%	If gIsShowLocal <> "N" Then	%>
									<TR>
										<TD WIDTH=100% COLSPAN = 4>
											<FIELDSET><LEGEND>�ΰ���</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS="TD5" NOWRAP>�ΰ�������</TD>
													<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="�ΰ�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenVatType()">&nbsp;<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="�ΰ�������"></TD>
													<TD CLASS="TD5" NOWRAP>�ΰ�����</TD>
													<TD CLASS="TD6" NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=OBJECT4 Name=txtVatRate style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 100px" title=FPDOUBLESINGLE ALT="�ΰ�����" tag="24"></OBJECT>');</SCRIPT>	&nbsp;%</TD>																				
												</TR>
												<TR>                    
													<TD CLASS=TD5 NOWRAP>�ΰ����ݾ�</TD>
													<TD CLASS=TD6 NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtVatAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="�ΰ����ݾ�" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
												    </TD>
													<TD CLASS=TD5 NOWRAP>�ΰ����ݾ�(�ڱ�)</TD>
													<TD CLASS=TD6 NOWRAP>									
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtVatLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="�ΰ����ݾ�(�ڱ�)" tag="24X2"> </OBJECT>');</SCRIPT> &nbsp;
		 												</TD>
												</TR>
												<TR>
													<TD CLASS="TD5" NOWRAP>�Ű�����</TD>
												    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportAreaCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="�Ű������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportAreaCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtReportAreaNm" SIZE=20 tag="24" ALT="�Ű������"></TD>
													<TD CLASS="TD5" NOWRAP>��꼭������</TD>																							    
													<TD CLASS="TD6" NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtIssuedDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��꼭������" tag="21X1" id=fpDateTime3> </OBJECT>');</SCRIPT>											    
													</TD>
												</TR>
											</TABLE>
											</FIELDSET>
										</TD>	
									</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtVatAmt"><INPUT TYPE=HIDDEN NAME="txtVatLocAmt"><INPUT TYPE=HIDDEN NAME="txtVatType"><INPUT TYPE=HIDDEN NAME="txtVatTypeNm">
<%	End If %>
									<TR>
										<TD WIDTH=100% HEIGHT= 100% VALIGN=TOP COLSPAN = 4>
											<TABLE <%=LR_SPACE_TYPE_60%>>							
												<TR>							
													<TD WIDTH="100%" HEIGHT=100% COLSPAN=4>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData HEIGHT="100%" tag="2" width="100%" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
											</TABLE>
										</TD>
									</TR>
								</TABLE>
								</DIV>
								<!-- �ι�° �� ����  -->
								<DIV ID="TabDiv" STYLE="DISPLAY: none;" SCROLL=no>
								<TABLE <%=LR_SPACE_TYPE_20%>>
									<TR>
										<TD WIDTH=100% HEIGHT= 100% VALIGN=TOP COLSPAN = 4>
											<TABLE <%=LR_SPACE_TYPE_60%>>							
												<TR>							
													<TD WIDTH="100%" HEIGHT=100% COLSPAN=4>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 HEIGHT="100%" tag="2" width="100%" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
													</TD>
												</TR>
											</TABLE>
										</TD>
									</TR>
								</TABLE>
								</DIV>			
							</TD>
							</TR>
						</TABLE>						
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR>
					<TD WIDTH=100% HEIGHT=10% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>							
							<TR>
								<TD CLASS="TD5" NOWRAP>���Ǹž�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="���Աݾ�" tag="24X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>���Ǹž�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="���Աݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle7></OBJECT>');</SCRIPT></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>
							<TR>
								<TD CLASS="TD5" NOWRAP>���Աݾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalRcptAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="���Աݾ�" tag="24X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>���Աݾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalRcptLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="���Աݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle7></OBJECT>');</SCRIPT></TD>
							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtTotalLocAmt"><INPUT TYPE=HIDDEN NAME="txtTotalRcptLocAmt">
<%	End If %>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=10>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2"	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows2"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadio"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hORGCHANGEID"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

