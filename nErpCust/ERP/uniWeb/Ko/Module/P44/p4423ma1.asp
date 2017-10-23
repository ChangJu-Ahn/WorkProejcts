<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4423ma1
'*  4. Program Name			: ���ְ����� ������ȸ 
'*  5. Program Desc			:
'*  6. Comproxy List		: +
'*  7. Modified date(First)	: 2001/11/27
'*  8. Modified date(Last) 	: 2003/05/26
'*  9. Modifier (First) 	: Jeon, jaehyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment		:
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit														'��: indicates that All variables must be declared in advance

Const BIZ_PGM_QRY1_ID = "p4423mb1.asp"								'��: �����Ͻ� ����(Qeury) ASP�� 
Const BIZ_PGM_QRY2_ID = "p4423mb2.asp"								'��: �����Ͻ� ����(Qeury) ASP�� 
'vspdData1
Dim C_BpCd				'= 1
Dim C_BpNm				'= 2
Dim C_CurCd				'= 3
Dim c_SubcontractAmt	'= 4
Dim C_TaxType			'= 5
Dim C_TaxTypeNm			'= 6
Dim C_TaxAmt			'= 7
Dim C_TotalCost			'= 8
Dim C_PlantCd			'= 9
Dim C_PlantNm			'= 10

'vspdData2
Dim C_ProdtOrderNo2		'= 1
Dim C_OrderQty2			'= 2
Dim C_OrderUnit2		'= 3
Dim C_ReportDt2			'= 4
Dim C_ResultQty2		'= 5
Dim C_CurCd2			'= 6
Dim C_SubcontractPrc2	'= 7
Dim c_SubcontractAmt2	'= 8
Dim C_TaxType2			'= 9
Dim C_TaxAmt2			'= 10
Dim C_TotalCost2		'= 11
Dim C_WcCd2				'= 12
Dim C_WcNm2				'= 13
Dim C_ItemCd2			'= 14
Dim C_ItemNm2			'= 15
Dim C_Spec2				'= 16
Dim C_TrackingNo2		'= 17
Dim C_OprNo2			'= 18
Dim C_Seq2				'= 19

Dim iDBSYSDate
Dim StartDate
Dim EndDate
Dim strYear
Dim strMonth
Dim strDay

iDBSYSDate = "<%=GetSvrDate%>"	
Call ExtractDateFrom(iDBSYSDate, parent.gServerDateFormat, parent.gServerDateType, StrYear, StrMonth, StrDay)
EndDate =  UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)			'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")			'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥	


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgStrPrevKey5
Dim lgStrPrevKey6
Dim lgStrPrevKey7
Dim lgLngCurRows
Dim lgSortKey1
Dim lgSortKey2

Dim lgLngCnt
Dim lgOldRow
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgStrPrevKey3 = ""                          'initializes Previous Key
    lgStrPrevKey4 = ""                          'initializes Previous Key
    lgStrPrevKey5 = ""                          'initializes Previous Key
    lgStrPrevKey6 = ""                          'initializes Previous Key
    lgStrPrevKey7 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'===========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromDt.Text = StartDate
	frm1.txtToDt.Text = EndDate
End Sub

'===========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ====================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'============================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	
	Call InitSpreadPosVariables(pvSpdNo)
	
		If pvSpdNo = "A" Or pvSpdNo = "*" Then 
			'------------------------------------------
			' Grid 1 - Operation Spread Setting
			'------------------------------------------
			With frm1.vspdData1
				 ggoSpread.Source = frm1.vspdData1
				 ggoSpread.Spreadinit "V20030602", , Parent.gAllowDragDropSpread
				.ReDraw = false
		
				.MaxCols = C_PlantNm + 1
				.MaxRows = 0
				
				Call GetSpreadColumnPos("A")
				
				ggoSpread.SSSetEdit		C_BpCd,			"����ó",	12
				ggoSpread.SSSetEdit		C_BpNm,			"����ó��",	20
				ggoSpread.SSSetEdit		C_CurCd,		"������ȭ", 10
				ggoSpread.SSSetFloat	C_SubcontractAmt,"���ֱݾ�", 15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
				ggoSpread.SSSetEdit		C_TaxType,		"VAT����",	15
				ggoSpread.SSSetEdit		C_TaxTypeNm,	"VAT����",	15
				ggoSpread.SSSetFloat	C_TaxAmt,		"VAT�ݾ�",	15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
				ggoSpread.SSSetFloat	C_TotalCost,	"�ѱݾ�",	15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
				ggoSpread.SSSetEdit		C_PlantCd,		"����",		12
				ggoSpread.SSSetEdit		C_PlantNm,		"�����",	20
				
				Call ggoSpread.SSSetColHidden( C_TaxType, C_TaxType, True)
				Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

				ggoSpread.SSSetSplit2(1)											'frozen ��� �߰� 
	
				.ReDraw = true

				Call SetSpreadLock("A") 

			End With
		End If	
		
		If pvSpdNo = "B" Or pvSpdNo = "*" Then 
			'------------------------------------------
			' Grid 2 - Operation Spread Setting
			'------------------------------------------
			With frm1.vspdData2
				 ggoSpread.Source = frm1.vspdData2
				 ggoSpread.Spreadinit "V20030602", , Parent.gAllowDragDropSpread
				.ReDraw = false
		
				.MaxCols = C_Seq2 + 1
				.MaxRows = 0
				
				Call GetSpreadColumnPos("B")
				ggoSpread.SSSetEdit		C_ProdtOrderNo2,	"������ȣ", 18
				ggoSpread.SSSetFloat	C_OrderQty2,		"��������", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
				ggoSpread.SSSetEdit		C_OrderUnit2,		"��������", 10
				ggoSpread.SSSetDate		C_ReportDt2,		"�԰���",	11, 2, parent.gDateFormat
				ggoSpread.SSSetFloat	C_ResultQty2,		"�԰����", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
				ggoSpread.SSSetEdit		C_CurCd2,			"������ȭ", 10
				ggoSpread.SSSetFloat	C_SubContractPrc2,	"���ִܰ�",15,"C" ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"		
				ggoSpread.SSSetFloat	C_SubcontractAmt2,	"���ֱݾ�", 15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
				ggoSpread.SSSetEdit		C_TaxType2,			"VAT����",	15
				ggoSpread.SSSetFloat	C_TaxAmt2,			"VAT�ݾ�",	15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
				ggoSpread.SSSetFloat	C_TotalCost2,		"�ѱݾ�",	15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
				ggoSpread.SSSetEdit		C_WcCd2,			"�۾���",	10
				ggoSpread.SSSetEdit		C_WcNm2,			"�۾����",	15
				ggoSpread.SSSetEdit		C_ItemCd2,			"ǰ��",		18
				ggoSpread.SSSetEdit		C_ItemNm2,			"ǰ���",	25
				ggoSpread.SSSetEdit		C_Spec2,			"�԰�",		25
				ggoSpread.SSSetEdit		C_TrackingNo2,		"Tracking No.", 25
				ggoSpread.SSSetEdit		C_OprNo2,			"����",		8
				ggoSpread.SSSetEdit		C_Seq2,				"����",		6
	
				Call ggoSpread.SSSetColHidden( C_OprNo2, C_OprNo2, True)
				Call ggoSpread.SSSetColHidden( C_Seq2, C_Seq2, True)
				Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)

				ggoSpread.SSSetSplit2(1)											'frozen ��� �߰� 
	
				.ReDraw = true

				Call SetSpreadLock("B") 

			End With
		End If
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	Select Case pvSpdNo
		Case "A"
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
		Case "B"	
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
	End Select			
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'vspdData1
		C_BpCd				= 1
		C_BpNm				= 2
		C_CurCd				= 3
		c_SubcontractAmt	= 4
		C_TaxType			= 5
		C_TaxTypeNm			= 6
		C_TaxAmt			= 7
		C_TotalCost			= 8
		C_PlantCd			= 9
		C_PlantNm			= 10
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		C_ProdtOrderNo2		= 1
		C_OrderQty2			= 2
		C_OrderUnit2		= 3
		C_ReportDt2			= 4
		C_ResultQty2		= 5
		C_CurCd2			= 6
		C_SubcontractPrc2	= 7
		c_SubcontractAmt2	= 8
		C_TaxType2			= 9
		C_TaxAmt2			= 10
		C_TotalCost2		= 11
		C_WcCd2				= 12
		C_WcNm2				= 13
		C_ItemCd2			= 14
		C_ItemNm2			= 15
		C_Spec2				= 16
		C_TrackingNo2		= 17
		C_OprNo2			= 18
		C_Seq2				= 19
	End If

End Sub
 
'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData1 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
			C_BpCd				= iCurColumnPos(1)
			C_BpNm				= iCurColumnPos(2)
			C_CurCd				= iCurColumnPos(3)
			c_SubcontractAmt	= iCurColumnPos(4)
			C_TaxType			= iCurColumnPos(5)
			C_TaxTypeNm			= iCurColumnPos(6)
			C_TaxAmt			= iCurColumnPos(7)
			C_TotalCost			= iCurColumnPos(8)
			C_PlantCd			= iCurColumnPos(9)
			C_PlantNm			= iCurColumnPos(10)
		Case "B"	
			ggoSpread.Source = frm1.vspdData2
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
			C_ProdtOrderNo2		= iCurColumnPos(1)
			C_OrderQty2			= iCurColumnPos(2)
			C_OrderUnit2		= iCurColumnPos(3)
			C_ReportDt2			= iCurColumnPos(4)
			C_ResultQty2		= iCurColumnPos(5)
			C_CurCd2			= iCurColumnPos(6)
			C_SubcontractPrc2	= iCurColumnPos(7)
			c_SubcontractAmt2	= iCurColumnPos(8)
			C_TaxType2			= iCurColumnPos(9)
			C_TaxAmt2			= iCurColumnPos(10)
			C_TotalCost2		= iCurColumnPos(11)
			C_WcCd2				= iCurColumnPos(12)
			C_WcNm2				= iCurColumnPos(13)
			C_ItemCd2			= iCurColumnPos(14)
			C_ItemNm2			= iCurColumnPos(15)
			C_Spec2				= iCurColumnPos(16)
			C_TrackingNo2		= iCurColumnPos(17)
			C_OprNo2			= iCurColumnPos(18)
			C_Seq2				= iCurColumnPos(19)

 	End Select
 
End Sub

'------------------------------------------ OpenBizPartner()  --------------------------------------------
'	Name : OpenBizparener()
'	Description : BpPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = frm1.txtBpCd.value 
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    arrField(2) = "BP_TYPE"
    arrField(3) = ""	
        
    arrHeader(0) = "BP"		
    arrHeader(1) = "BP��"		
    arrHeader(2) = "Bp ����"		
    arrHeader(3) = ""
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtBpCd.Value    = arrRet(0)		
		frm1.txtBpNm.Value    = arrRet(1)	
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBpCd.focus
	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)		' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
    
    arrHeader(0) = "����"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
		
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenConWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & " And INSIDE_FLG = " & FilterVar("N", "''", "S") & " " 'Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtWCCd.Value    = arrRet(0)		
		frm1.txtWCNm.Value    = arrRet(1)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
		
End Function

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================== 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    Call InitSpreadSheet("*")                                               '��: Setup the Spread sheet
	Call SetDefaultVal
    Call InitVariables														'��: Initializes local global variables
	
	Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	End If
	
	frm1.txtBpCd.focus 
	Set gActiveElement = document.activeElement
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    '----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
 		
 		lgOldRow = Row
		
		lgStrPrevKey2 = ""
			
		frm1.vspdData2.MaxRows = 0
			  		
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
 		
	Else
 		'------ Developer Coding part (Start)
 		If lgOldRow <> Row Then
			
			frm1.vspdData1.Col = 1
			frm1.vspdData1.Row = row
	
			lgOldRow = Row
			
			lgStrPrevKey2 = ""
				
			frm1.vspdData2.MaxRows = 0
			  		
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
			
		End If
	 	'------ Developer Coding part (End)	
 	End If
	
End Sub

Sub vspdData2_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
    
 	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData2 
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey2 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
 			lgSortKey2 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
 	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'==========================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'==========================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )
	On Error Resume Next	
End Sub

Sub vspdData2_Change(ByVal Col , ByVal Row )
	On Error Resume Next
End Sub

'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData1_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_SubcontractAmt
            Call EditModeCheck(frm1.vspdData1, Row, C_CurCd, C_SubcontractAmt, "A" ,"I", Mode, "X", "X") 
        Case C_TaxAmt
            Call EditModeCheck(frm1.vspdData1, Row, C_CurCd, C_TaxAmt, "A" ,"I", Mode, "X", "X")  
        Case C_TotalCost
            Call EditModeCheck(frm1.vspdData1, Row, C_CurCd, C_TotalCost, "A" ,"I", Mode, "X", "X")         
    End Select
End Sub

Sub vspdData2_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_SubcontractPrc2
            Call EditModeCheck(frm1.vspdData2, Row, C_CurCd2, C_SubcontractPrc2, "C" ,"I", Mode, "X", "X") 
        Case C_SubcontractAmt2
            Call EditModeCheck(frm1.vspdData2, Row, C_CurCd2, C_SubcontractAmt2, "A" ,"I", Mode, "X", "X") 
        Case C_TaxAmt2
            Call EditModeCheck(frm1.vspdData2, Row, C_CurCd2, C_TaxAmt2, "A" ,"I", Mode, "X", "X")  
        Case C_TotalCost2
            Call EditModeCheck(frm1.vspdData2, Row, C_CurCd2, C_TotalCost2, "A" ,"I", Mode, "X", "X")         
    End Select
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData1 

    If Row >= NewRow Then
        Exit Sub
    End If
	'----------  Coding part  -------------------------------------------------------------
    End With

End Sub

Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData2

    If Row >= NewRow Then
        Exit Sub
    End If
	'----------  Coding part  -------------------------------------------------------------
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop)  Then
		If lgStrPrevKey5 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			
			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("B")
End Sub  

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    Dim IntRetCD 

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = "" 
	End If
	
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtWCCd.value = "" Then
		frm1.txtWCNm.value = "" 
	End If
	

	If ValidDateCheck(frm1.txtFromDt, frm1.txtTODt) = False Then Exit Function	
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then											'��: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function															'��: Query db data
	End If

    FncQuery = True																'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()															'��: Protect system from crashing
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                     '��:ȭ�� ����, Tab ���� 
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

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow                   
    Dim StrNextKey      
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                           '��: Protect system from crashing

	Dim strVal
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001		'��: 
			strVal = strVal & "&txtBpCd=" & Trim(.hBpCd.value)
			strVal = strVal & "&txtFromDt=" & Trim(.hFromDt.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToDt=" & Trim(.hToDt.value)
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
			strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001		'��:
			strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value) 
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.text)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.value)
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
			strVal = strVal & "&lgStrPrevKey4=" & lgStrPrevKey4
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
	Call SetToolBar("11000000000111")										'��: ��ư ���� ���� 
	lgOldRow = 1
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
    
    lgIntFlgMode = parent.OPMD_UMODE										'��: Indicates that current mode is Update mode
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbDtlQuery() 
    Dim strVal
    Dim strBpCd, strPlantCd, strCurCd, strTaxType   
    
	frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
	
	frm1.vspdData1.Col = C_BpCd
	strBpCd = Trim(frm1.vspdData1.Text)
	
	frm1.vspdData1.Col = C_PlantCd
	strPlantCd = Trim(frm1.vspdData1.Text)
	
	frm1.vspdData1.Col = C_CurCd
	strCurCd = Trim(frm1.vspdData1.Text)
	
	frm1.vspdData1.Col = C_TaxType
	strTaxType = Trim(frm1.vspdData1.Text)
	
    DbDtlQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
        
    With frm1
    
		Call LayerShowHide(1)
		
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtBpCd=" & strBpCd
			strVal = strVal & "&txtFromDt=" & Trim(.hFromDt.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToDt=" & Trim(.hToDt.value)
			strVal = strVal & "&txtPlantCd=" & strPlantCd					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
			strVal = strVal & "&txtCurCd=" & strCurCd
			strVal = strVal & "&txtTaxType=" & strTaxType
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&lgStrPrevKey5=" & lgStrPrevKey5
			strVal = strVal & "&lgStrPrevKey6=" & lgStrPrevKey6
			strVal = strVal & "&lgStrPrevKey7=" & lgStrPrevKey7
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
		Else
					
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtBpCd=" & strBpCd 
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.text)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPlantCd=" & strPlantCd					'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.value)
			strVal = strVal & "&txtCurCd=" & strCurCd
			strVal = strVal & "&txtTaxType=" & strTaxType
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&lgStrPrevKey5=" & lgStrPrevKey5
			strVal = strVal & "&lgStrPrevKey6=" & lgStrPrevKey6
			strVal = strVal & "&lgStrPrevKey7=" & lgStrPrevKey7
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbDtlQuery = True

End Function

Function DbDtlQueryOk()														'��: ��ȸ ������ ������� 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
	On Error Resume Next
End Function


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
Sub PopRestoreSpreadColumnInf()
Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.ID)
	Call ggoSpread.ReOrderingSpreadData()
    
End Sub 


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ְ����񳻿���ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
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
									<TD CLASS=TD5 NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizPartner() ">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="����ó��"></TD>
									<TD CLASS=TD5 NOWRAP>�԰�����</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/p4423ma1_I361231869_txtFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p4423ma1_I753024446_txtToDt.js'></script>
									</TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="11XXXU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConWC()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=25 tag="14" ALT="�۾����"></TD>
								</TR>															
							</TABLE>							
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="50%">
									<script language =javascript src='./js/p4423ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="50%">
									<script language =javascript src='./js/p4423ma1_B_vspdData2.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBpCd" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
