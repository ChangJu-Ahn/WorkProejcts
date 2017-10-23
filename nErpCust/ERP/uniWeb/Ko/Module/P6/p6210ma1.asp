<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p6210ma1
'*  4. Program Name			: ����������� 
'*  5. Program Desc			:
'*  6. Comproxy List		: 
'*  7. Modified date(First)	: 2005/10/21
'*  8. Modified date(Last) 	: 2005/10/21
'*  9. Modifier (First) 	: Chen, Jae Hyun
'* 10. Modifier (Last)		: Chen, Jae Hyun
'* 11. Comment				:
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'      Park Kye Jin         : ������ȹ����/�Ϸ��ȹ����/�ǿϷ��� ����(2003.04.07)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************** -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->							<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ��� -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "p6210mb1.asp"								'��: �����Ͻ� ����(Qeury) ASP�� 
Const BIZ_PGM_SAVE_ID	= "p6210mb2.asp"								'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Dim C_Select
Dim C_CastRsltFlag	
Dim C_DelFlag			
Dim C_ProdtOrderNo
Dim C_OprNo	
Dim C_Seq	
Dim C_ShiftCd			
Dim C_ItemCd				
Dim C_ItemNm				
Dim C_Spec	
Dim C_ReportDt
Dim C_ProdQty				
Dim C_ProdtOrderUnit
Dim C_FacilityCd
Dim C_FacilityPopup
Dim C_FacilityNm
Dim C_CastCd
Dim C_CastPopup
Dim C_CastNm
Dim C_CurCount
Dim C_Cavi
Dim C_InputQty
Dim C_Remark				
Dim C_WcCd	
Dim C_WcNm
Dim C_TrackingNo
Dim C_ItemGroupCd
Dim C_ItemGroupNm


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgSortKey 
Dim lgEditUndoKey

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  -------------------------------------------------------------- 
Dim IsOpenPop          
Dim lgButtonSelection
Dim lgRedrewFlg

Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)		'���ιٲ� ��¥ ����											
StartDate = UNIDateAdd("D",-10,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ó�� ��¥ 
EndDate = UNIDateAdd("D", 20,LocSvrDate, parent.gDateFormat)	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call AppendNumberPlace("6", "11", "0")
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
		 
    Call ggoOper.LockField(Document, "Q")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet

	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables
    Call InitComboBox
    
	Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Ucase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtFromItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
    
End Sub

'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey1 = ""                          'initializes Previous Key1
    lgStrPrevKey2 = ""                          'initializes Previous Key2
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgEditUndoKey = False
   	lgButtonSelection = "DESELECT"
	'frm1.btnAutoSel.disabled = True
	'frm1.btnAutoSel.value = "��ü����"
	    
End Sub

'==========================================  2.2.6 InitData()  ========================================== 
'	Name : InitData()
'	Description : Combo Display
'======================================================================================================== 
Sub InitData(ByVal lngStartRow)
	Dim intRow
	Dim intIndex
	Dim strDelFlag
	
	frm1.vspdData.Redraw = False
	
	With frm1.vspdData
		For intRow = lngStartRow To .MaxRows
			Call .GetText(C_DelFlag, intRow, strDelFlag)
			.Row = intRow
			.Col = C_CastRsltFlag

			If Trim(.Text) = "Y" Then
				
				ggoSpread.SpreadUnLock		C_FacilityCd, intRow , C_FacilityCd ,intRow
				ggoSpread.SpreadUnLock		C_FacilityPopup, intRow , C_FacilityPopup ,intRow
				ggoSpread.SpreadUnLock		C_CastCd, intRow , C_CastCd ,intRow
				ggoSpread.SpreadUnLock		C_CastPopup, intRow , C_CastPopup ,intRow
				ggoSpread.SSSetRequired		C_CastCd,			intRow, intRow
				ggoSpread.SpreadUnLock		C_InputQty, intRow , C_InputQty ,intRow
				ggoSpread.SSSetRequired		C_InputQty,			intRow, intRow
				ggoSpread.SpreadUnLock		C_Remark, intRow , C_Remark ,intRow
				
				If strDelFlag = "Y" Then
					.Col = C_ProdtOrderNo
					.ForeColor = vbRed
				End If
			Else
				ggoSpread.SSSetProtected	C_FacilityCd,			intRow, intRow
				ggoSpread.SSSetProtected	C_FacilityPopup,		intRow, intRow
				ggoSpread.SpreadLock		C_FacilityCd,			intRow , C_FacilityCd,		intRow
				ggoSpread.SpreadLock		C_FacilityPopup,		intRow , C_FacilityPopup,	 intRow	
				ggoSpread.SSSetProtected	C_CastCd,				intRow, intRow
				ggoSpread.SSSetProtected	C_CastPopup,			intRow, intRow
				ggoSpread.SpreadLock		C_CastCd,				intRow , C_CastCd,		intRow
				ggoSpread.SpreadLock		C_CastPopup,			intRow , C_CastPopup,	 intRow
				ggoSpread.SSSetProtected	C_InputQty,				intRow, intRow
				ggoSpread.SpreadLock		C_InputQty,				intRow , C_Remark,	 intRow
				ggoSpread.SSSetProtected	C_Remark,				intRow, intRow
				ggoSpread.SpreadLock		C_Remark,				intRow , C_Remark,	 intRow
				
				If strDelFlag = "Y" Then
					ggoSpread.SSSetProtected	C_Select,			intRow, intRow
					ggoSpread.SpreadLock		C_Select,			intRow , C_Select,		intRow
				End If
				
			End If	
			
		Next	
	End With
	
	frm1.vspdData.Redraw = True
End Sub

'========================================  2.1.3 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : ComboBox�� ����Ÿ Setting
'====================================================================================================
Sub InitComboBox()

	Call SetCombo(frm1.cboCastRsltFlg, "N", "�̵��")
    Call SetCombo(frm1.cboCastRsltFlg, "Y", "���")		'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
    
End Sub    

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromDt.text = StartDate
    frm1.txtToDt.text   = EndDate
	'frm1.btnAutoSel.disabled = True
	'frm1.btnAutoSel.value = "��ü����"
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'================================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

    With frm1.vspdData
    .ReDraw = false

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20050815", , Parent.gAllowDragDropSpread
	
	.MaxCols = C_ItemGroupNm + 1
	.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetCheck	C_Select, "��������", 8,,,1
	ggoSpread.SSSetEdit		C_CastRsltFlag, "��������", 6
	ggoSpread.SSSetEdit		C_DelFlag, "��������", 6
	ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 18
	ggoSpread.SSSetEdit		C_OprNo, "����", 6
	ggoSpread.SSSetEdit		C_Seq, "����", 6
	ggoSpread.SSSetEdit		C_ShiftCd, "Shift", 8
	ggoSpread.SSSetEdit		C_ItemCd, "ǰ��", 18
	ggoSpread.SSSetEdit		C_ItemNm, "ǰ���", 25
	ggoSpread.SSSetEdit		C_Spec, "�԰�", 25
	ggoSpread.SSSetDate		C_ReportDt, "������", 10, 2, parent.gDateFormat
	ggoSpread.SSSetFloat	C_ProdQty, "���귮", 15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_ProdtOrderUnit, "��������", 8
	ggoSpread.SSSetEdit		C_FacilityCd, "�����ڵ�", 12,,,18,2
	ggoSpread.SSSetButton 	C_FacilityPopup
	ggoSpread.SSSetEdit		C_FacilityNm, "�����", 20
	ggoSpread.SSSetEdit		C_CastCd, "�����ڵ�", 12,,,18,2
	ggoSpread.SSSetButton 	C_CastPopup
	ggoSpread.SSSetEdit		C_CastNm, "�����ڵ��", 20
	ggoSpread.SSSetFloat	C_CurCount, "����Ÿ��", 15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_Cavi, "Ÿ�������", 15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetFloat	C_InputQty, "�ݿ�����", 15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	ggoSpread.SSSetEdit		C_Remark, "���", 30,,,25
	ggoSpread.SSSetEdit		C_WcCd, "�۾���", 10
	ggoSpread.SSSetEdit		C_WcNm, "�۾����", 20
	ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.", 25
	ggoSpread.SSSetEdit 	C_ItemGroupCd, "ǰ��׷�",	15
	ggoSpread.SSSetEdit		C_ItemGroupNm, "ǰ��׷��", 30

	Call ggoSpread.MakePairsColumn(C_FacilityCd, C_FacilityPopup)
	Call ggoSpread.SSSetColHidden(C_CastRsltFlag, C_DelFlag , True)
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols , True)
	ggoSpread.SSSetSplit2(5)											'frozen ��� �߰� 
	
	.ReDraw = true

	Call SetSpreadLock

    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_ProdtOrderNo, -1, C_ItemGroupNm
	.vspdData.ReDraw = True
	
	End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
  On Error Resume Next
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_Select			= 1
	C_CastRsltFlag		= 2
	C_DelFlag			= 3
	C_ProdtOrderNo		= 4
	C_OprNo				= 5
	C_Seq				= 6
	C_ShiftCd			= 7
	C_ItemCd			= 8		
	C_ItemNm			= 9	
	C_Spec				= 10 
	C_ReportDt			= 11
	C_ProdQty			= 12	
	C_ProdtOrderUnit	= 13
	C_FacilityCd		= 14
	C_FacilityPopup		= 15
	C_FacilityNm		= 16
	C_CastCd			= 17
	C_CastPopup			= 18
	C_CastNm			= 19
	C_CurCount			= 20
	C_Cavi				= 21
	C_InputQty			= 22
	C_Remark			= 23
	C_WcCd				= 24
	C_WcNm				= 25
	C_TrackingNo		= 26
	C_ItemGroupCd		= 27
	C_ItemGroupNm		= 28

End Sub
 

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
  	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
  	Case "A"
 		ggoSpread.Source = frm1.vspdData 
  		
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
  		
		C_Select			= iCurColumnPos(1)
		C_CastRsltFlag		= iCurColumnPos(2)
		C_DelFlag			= iCurColumnPos(3)		
		C_ProdtOrderNo		= iCurColumnPos(4)
		C_OprNo				= iCurColumnPos(5)
		C_Seq				= iCurColumnPos(6)
		C_ShiftCd			= iCurColumnPos(7)
		C_ItemCd			= iCurColumnPos(8)		
		C_ItemNm			= iCurColumnPos(9)	
		C_Spec				= iCurColumnPos(10)
		C_ReportDt			= iCurColumnPos(11)
		C_ProdQty			= iCurColumnPos(12)	
		C_ProdtOrderUnit	= iCurColumnPos(13)
		C_FacilityCd		= iCurColumnPos(14)
		C_FacilityPopup		= iCurColumnPos(15)
		C_FacilityNm		= iCurColumnPos(16)
		C_CastCd			= iCurColumnPos(17)
		C_CastPopup			= iCurColumnPos(18)
		C_CastNm			= iCurColumnPos(19)
		C_CurCount			= iCurColumnPos(20)
		C_Cavi				= iCurColumnPos(21)
		C_InputQty			= iCurColumnPos(22)
		C_Remark			= iCurColumnPos(23)	
		C_WcCd				= iCurColumnPos(24)
		C_WcNm				= iCurColumnPos(25)
		C_TrackingNo		= iCurColumnPos(26)
		C_ItemGroupCd		= iCurColumnPos(27)
		C_ItemGroupNm		= iCurColumnPos(28)
  	End Select
  
End Sub

'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++

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
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo(Byval strCode, Byval strWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field��(0)
	arrField(1) = 2 '"ITEM_NM"					' Field��(1)
    
    iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet, strWhere)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtFromItemCd.focus

End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'----------------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "ST"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtFromItemCd.value)
	arrParam(8) = ""
	
	iCalledAspName = AskPRAspName("p4111pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "ǰ��׷�"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "ǰ��׷�"
	arrHeader(1) = "ǰ��׷��"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function

'--------------------------------------  OpenTrackingInfo()  ---------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode, Byval iWhere)
    If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""	
	
	iCalledAspName = AskPRAspName("p4600pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4600pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
		
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
		Call DisplayMsgBox("971012","X", "����","X")
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 			' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
	
End Function

'------------------------------------------  OpenFacilityCd()  -------------------------------------------------
'	Name : OpenFacilityCd()
'	Description : Condition Work Center PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenFacilityCd(ByVal strCode, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"											' �˾� ��Ī 
	arrParam(1) = " Y_FACILITY "											' TABLE ��Ī 
	arrParam(2) = Trim(strCode)									' Code Condition
	arrParam(3) = ""				'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = " SET_PLANT = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 	&	_	
					" AND USE_YN = " &	FilterVar("Y", "''", "S")													' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "FACILITY_CD"													' Field��(0)
    arrField(1) = "FACILITY_NM"													' Field��(1)
    
    arrHeader(0) = "�����ڵ�"												' Header��(0)
    arrHeader(1) = "�����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetFacilityCd(arrRet, Row)
	Else
		Exit Function
	End If	
	
End Function

'------------------------------------------  OpenCastCd()  -------------------------------------------------
'	Name : OpenCastCd()
'	Description : Condition Work Center PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenCastCd(ByVal strCode, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	Dim strWCCd, strItemCd
	
	Dim strWhere, ArrWhere
	Dim pvCnt

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	With frm1.vspdData
		.Col = C_WcCd
		.Row = Row
		strWcCd = Trim(.Text)
		.Col = C_ItemCd
		strItemCd = Trim(.Text)
	End With
	
	strWhere = " AND ( "
	
	Redim ArrWhere(9)
	
	For pvCnt =  0 To 9
		ArrWhere(pvCnt) = " item_cd_" & Cstr(1 + pvCnt) & " = " & Filtervar(strItemCd, "''", "S") 
	Next
	
	strWhere = strWhere & Join(ArrWhere, " Or ") & ") "
	
	arrParam(0) = "�����˾�"											' �˾� ��Ī 
	arrParam(1) = " Y_CAST "											' TABLE ��Ī 
	arrParam(2) = Trim(strCode)									' Code Condition
	arrParam(3) = ""				'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = " SET_PLANT = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 	&	_	
					" AND USE_YN = " &	FilterVar("Y", "''", "S")	 & _
					strWhere												' Where Condition
					'" AND SET_PLACE = " & FilterVar(strWCCd, "''", "S") & _
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "CAST_CD"													' Field��(0)
    arrField(1) = "CAST_NM"													' Field��(1)
    
    arrHeader(0) = "�����ڵ�"											' Header��(0)
    arrHeader(1) = "�����ڵ��"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCastCd(arrRet, Row)
	Else
		Exit Function
	End If	
	
End Function


'------------------------------------------  OpenOprCd()  -------------------------------------------------
'	Name : OpenOprCd()
'	Description : Condition Operation PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprCd()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtOprCd.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtProdOrderNo.value = "" Then
		Call displaymsgbox("971012","X", "����������ȣ","X")
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4112PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4112PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdOrderNo.value
	arrParam(2) = "Y"
	

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetOprCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprCd.focus
	
End Function
'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'-----------------------------------------------------------------------------------------------------------
Function OpenOprRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Function
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		arrParam(1) = .Text
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4111ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetConPlant()  -----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(byval arrRet, byval strWhere)
	Select Case Trim(strWhere)
		Case "1"
			frm1.txtFromItemCd.Value    = arrRet(0)		
			frm1.txtFromItemNm.Value    = arrRet(1)	
		
		Case "2"
			frm1.txtToItemCd.Value    = arrRet(0)		
			frm1.txtToItemNm.Value    = arrRet(1)	
		
	End Select	
		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)
End Function

'=========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetTrackingNo()  -------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function

'------------------------------------------  SetConWC()  -------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConWC(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetFacilityCd()  ----------------------------------------------
'	Name : SetFacilityCd()
'	Description : Work Center Popup for Grid 2 ���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetFacilityCd(Byval arrRet, Byval Row)
	With frm1
	    .vspdData.Row = Row
		.vspdData.Col = C_FacilityCd
		.vspdData.Text = UCase(arrRet(0))
		.vspdData.Col = C_FacilityNm
		.vspdData.Text = UCase(arrRet(1))
		Call vspdData_Change(C_FacilityCd, .vspdData.Row)
	End With
	
End Function


'------------------------------------------  SetCastCd()  ----------------------------------------------
'	Name : SetCastCd()
'	Description : Work Center Popup for Grid 2 ���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCastCd(Byval arrRet, Byval Row)
	With frm1
	    .vspdData.Row = Row
		.vspdData.Col = C_CastCd
		.vspdData.Text = UCase(arrRet(0))
		.vspdData.Col = C_CastNm
		.vspdData.Text = UCase(arrRet(1))
		Call vspdData_Change(C_CastCd, .vspdData.Row)
	End With
	
End Function



'------------------------------------------  SetOprCd()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprCd(byval arrRet)
	frm1.txtOprCd.Value    = arrRet(0)		
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'**********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

Function btnAutoSel_onClick()

	lgRedrewFlg = False

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "��ü����"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "��ü�������"
	End If

	Dim index,Count
	Dim strFlag
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Select
		
		strFlag = frm1.vspdData.Value
		
		If lgButtonSelection = "SELECT" Then
			frm1.vspdData.Value = 1
			frm1.vspdData.Col = 0 
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData.Value = 0
			frm1.vspdData.Col = 0 
			'ggoSpread.SSDeleteFlag Index
			frm1.vspdData.Text=""
		End if

	Next 
	
	frm1.vspdData.ReDraw = true

	lgRedrewFlg = True

End Function

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *************************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
  	
  	If lgIntFlgMode = Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")         'ȭ�麰 ���� 
	Else
		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
	End If
  	
  	gMouseClickStatus = "SPC"   
     
  	Set gActiveSpdSheet = frm1.vspdData
     
  	If frm1.vspdData.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData 
  		If lgSortKey = 1 Then
  			ggoSpread.SSSort Col					'Sort in Ascending
  			lgSortKey = 2
  		Else
  			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
  			lgSortKey = 1
  		End If
 	
  	End If
  	
  	'------ Developer Coding part (Start)
  	
 	'------ Developer Coding part (End)

End Sub
 
'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
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
 
 	If NewCol = C_Select or Col = C_Select Then
 		Cancel = True
 		Exit Sub
 	End If
 
     ggoSpread.Source = frm1.vspdData
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("A")
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
Sub PopRestoreSpreadColumnInf()
      ggoSpread.Source = gActiveSpdSheet
     Call ggoSpread.RestoreSpreadInf()
     Call InitSpreadSheet
     Call InitData(1)
     Call ggoSpread.ReOrderingSpreadData
End Sub 

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
    Dim pvSelect
    
    With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
			If Row < 1 Then Exit Sub
		Select Case Col
			Case C_Select
			
				If lgRedrewFlg = True Then .ReDraw = false
				
				.Row = Row
				.Col = C_CastRsltFlag
				pvSelect = .Text
				
				If ButtonDown = 1 Then
					If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
						Exit Sub
					End If
					
					If lgEditUndoKey = True Then
						lgEditUndoKey = False
						Exit Sub
					End If
						
					ggoSpread.UpdateRow Row
					ggoSpread.SpreadUnLock		C_FacilityCd, Row , C_FacilityCd ,Row
					ggoSpread.SpreadUnLock		C_FacilityPopup, Row , C_FacilityPopup ,Row
					ggoSpread.SpreadUnLock		C_Remark, Row , C_Remark ,Row
					ggoSpread.SpreadUnLock		C_CastCd, Row , C_CastPopup ,Row
					ggoSpread.SSSetRequired		C_CastCd,			Row, Row
					ggoSpread.SpreadUnLock		C_InputQty, Row , C_InputQty ,Row
					ggoSpread.SSSetRequired		C_InputQty,			Row, Row
					
					Call LookUpTopCastCd(Row)
						
				Else
					If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
						Exit Sub
					End If
					
					If lgEditUndoKey = True Then
						lgEditUndoKey = False
						Exit Sub
					End If
					
					If pvSelect = "Y" Then
						ggoSpread.UpdateRow Row
					Else
						ggoSpread.SSDeleteFlag Row,Row
					End If
					
					Call .SetText(C_FacilityCd, Row, "")
					Call .SetText(C_FacilityNm, Row, "")
					Call .SetText(C_CastCd, Row, "")
					Call .SetText(C_CastNm, Row, "")
					Call .SetText(C_Remark, Row, "")
					Call .SetText(C_Cavi, Row, "0")
					Call .SetText(C_CurCount, Row, "0")
					Call .SetText(C_InputQty, Row, "0")
					
					ggoSpread.SSSetProtected C_FacilityCd,			Row, Row
					ggoSpread.SSSetProtected C_FacilityPopup,		Row, Row
					ggoSpread.SSSetProtected C_Remark,			Row, Row
					ggoSpread.SpreadLock	C_FacilityCd,		Row , C_FacilityCd,		Row
					ggoSpread.SpreadLock	C_FacilityPopup,	Row , C_FacilityPopup,	 Row
					ggoSpread.SpreadLock	C_Remark,			Row , C_Remark,	 Row	
					ggoSpread.SSSetProtected C_CastCd,			Row, Row
					ggoSpread.SSSetProtected C_CastPopup,		Row, Row
					ggoSpread.SpreadLock	C_CastCd,		Row , C_CastCd,		Row
					ggoSpread.SpreadLock	C_CastPopup,	Row , C_CastPopup,	 Row
					ggoSpread.SSSetProtected C_InputQty,			Row, Row
					ggoSpread.SpreadLock	C_InputQty,	Row , C_InputQty,	 Row		
				End If		

				If lgRedrewFlg = True Then .ReDraw = True
			
			Case C_FacilityPopup
				.Col = C_FacilityCd
				.Row = Row
				Call OpenFacilityCd(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_FacilityCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement
				
			Case C_CastPopup
				.Col = C_CastCd
				.Row = Row
				Call OpenCastCd(.Text, Row)
				Call SetActiveCell(frm1.vspdData,C_CastCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement
					
		End Select
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData 

    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------  

    End With

End Sub

'=======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)         
	
	Dim strCastCd
	
	With frm1.vspdData

		.Row = Row

		Select Case Col
		
			Case C_InputQty, C_FacilityCd, C_Remark
		
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
				
			Case C_CastCd
				
				Call .GetText(C_CastCd, Row, strCastCd)
				Call LookUpCastCd(strCastCd, Row)
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow Row
				
		End Select

	End With
	

End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
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
'   Event Name : txtRcptDT_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtRcptDT_DblClick(Button)
    If Button = 1 Then
        frm1.txtRcptDT.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtRcptDT.Focus
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


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False                                            '��: Processing is NG

    Err.Clear                                                   '��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    
    Call InitVariables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then							'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If

    FncQuery = True												'��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    On Error Resume Next                                                   '��: Protect system from crashing    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    
    FncSave = False												'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing
   
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then				'��: Check required field(Multi area)
       Exit Function
    End If
        
    '-----------------------
    'Save function call area
    '-----------------------
    
    If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True												'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function	 
    ggoSpread.Source = frm1.vspdData
    lgEditUndoKey = True	
    ggoSpread.EditUndo                                             '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 

End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                                  '��: Protect system from crashing
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                             '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                             '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
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

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'**********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear

	Dim strVal
    
    With frm1
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&txtFromItemCd=" & Trim(.hFromItemCd.value)
		strVal = strVal & "&txtToItemCd=" & Trim(.hToItemCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)
		strVal = strVal & "&txtOprCD=" & Trim(.hOprCD.value)
		strVal = strVal & "&txtWcCd=" & Trim(.hWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtFromDt=" & Trim(.hFromDt.value)
		strVal = strVal & "&txtToDt=" & Trim(.hToDt.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		strVal = strVal & "&cboResultFlg=" & Trim(.hResultFlg.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows 
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtFromItemCd=" & Trim(.txtFromItemCd.value)
		strVal = strVal & "&txtToItemCd=" & Trim(.txtToItemCd.value)
		strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)
		strVal = strVal & "&txtOprCD=" & Trim(.txtOprCD.value)
		strVal = strVal & "&txtWcCd=" & Trim(.txtWcCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
		strVal = strVal & "&cboResultFlg="  & Trim(.cboCastRsltFlg.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	End If 
	
    Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)
	
	Dim LngRow
	
	Call SetToolbar("11001001000111")

    Call ggoOper.LockField(Document, "N")
	
	'frm1.btnAutoSel.disabled = False
    '-----------------------
    'Reset variables area
    '-----------------------
    Call InitData(LngMaxRow)
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
	
End Function

Function DbQueryNotOk()														'��: ��ȸ ������ ������� 
	Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	
End Function	
'========================================================================================
' Function Name : DbSave
' Function Desc : This function is to execute transaction.
'========================================================================================
Function DbSave() 

    Dim lRow    
	Dim strVal
	
	Dim tmpCastRsltFlag, tmpSelectFlag
	
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size
	
	DbSave = False                                                          '��: Processing is NG
    
    Call LayerShowHide(1)
    
    frm1.txtMode.value = parent.UID_M0002
	frm1.txtUpdtUserId.value = parent.gUsrID
	frm1.txtInsrtUserId.value = parent.gUsrID
		
	'-----------------------
	'Data manipulate area
	'-----------------------
	iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
    
	'102399byte
	iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
	'������ �ʱ�ȭ 
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)		

	iTmpCUBufferCount = -1
	
	strCUTotalvalLen = 0

	With frm1.vspdData
	
		For lRow = 1 To .MaxRows
		
		    .Row = lRow
		    .Col = 0
		    
		    If .Text = ggoSpread.UpdateFlag Then
				
				.Col = C_CastRsltFlag
				If Trim(.Text) = "Y" Then
					tmpCastRsltFlag = "Y"
				Else
					tmpCastRsltFlag = "N"
				End If
				
				.Col = C_Select
				If Trim(.Text) = "1" Then
					tmpSelectFlag = "Y"
				Else
					tmpSelectFlag = "N"
				End If
				
				If tmpCastRsltFlag = "Y" AND tmpSelectFlag = "Y" Then
					strVal = "Update" & iColSep
				ElseIf tmpCastRsltFlag = "Y" AND tmpSelectFlag = "N"  Then
					strVal = "Delete" & iColSep
				ElseIf tmpCastRsltFlag = "N" AND tmpSelectFlag = "Y" Then
					strVal = "Create" & iColSep
				Else
					strVal = ""
				End If

				
				If Not( tmpCastRsltFlag = "N" AND tmpSelectFlag = "N" ) Then
						
					'//Ref. ConstBas\Y0\BCY623_PMngCastRslt.bas
					.Col = C_ProdtOrderNo			
					strVal = strVal & UCase(Trim(.Text)) & iColSep	'ProdtOrderNo
					.Col = C_OprNo					
					strVal = strVal & Trim(.Text) & iColSep	'OprNo
					.Col = C_Seq					
					strVal = strVal & CInt(Trim(.Text)) & iColSep			'Seq
					.Col = C_CastCd	
					strVal = strVal & UCase(Trim(.Text)) & iColSep					'cast_cd
					.Col = C_FacilityCd	
					strVal = strVal & UCase(Trim(.Text)) & iColSep					'facility_cd
					.Col = C_InputQty
					If UNICDbl(.Text) = 0  and tmpSelectFlag  = "Y" Then
						Call DisplayMsgBox("970022", "X", "�ݿ�����", "0")	
						Set gActiveElement = document.activeElement 
						Call LayerShowHide(0)
						Call RemovedivTextArea()
						Exit Function
					End If
					strVal = strVal & UNIConvNum(Trim(.Text), 1) & iColSep					'facility_cd
					.Col = C_Remark	
					strVal = strVal & Trim(.Text) & iColSep					'remark

					strVal = strVal & lRow & iRowSep						'Count (to trace error row)
							
					If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
					                            
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
				End If	    
						
			End If	
			            
		Next
		
		If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)
		End If   
		
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)							'��: �����Ͻ� ASP �� ���� 
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()															'��: ���� ������ ���� ���� 
   
    Call InitVariables
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call RemovedivTextArea
    Call MainQuery()

End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

End Function


'-------------------------------------  LookUpItem ByPlant()  -----------------------------------------
'	Name : LookUpItem ByPlant()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 
Function LookUpCastCd(Byval StrCastCd, Byval Row)
    
	Dim strVal
	Dim strSelect, strWhere
	Dim lgF0
	Dim pvCnt
	Dim tmpArrSQL(9)
	Dim strDelFlag, strItemCd, strWcCd
	Dim gComNum1000, gComNumDec, gAPNum1000, gAPNumDec
	
	Dim strInputQty, strProdQty, strCaviQty
	
	gComNum1000 = parent.gComNum1000
	gComNumDec = parent.gComNumDec
	gAPNum1000 = parent.gAPNum1000
	gAPNumDec = parent.gAPNumDec

	If StrCastCd = "" Then Exit Function
	
	frm1.vspdData.Col = C_CastCd
	frm1.vspdData.Row = Row		
	
	Call frm1.vspdData.GetText(C_ItemCd, Row, strItemCd)
	
	strSelect = " CAST_NM, CUR_ACCNT, PRS_UNIT  "
	strWhere = " SET_PLANT = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	strWhere = strWhere & " AND USE_YN = 'Y' AND CAST_CD = " & FilterVar(frm1.vspdData.Text, "''", "S")
	strWhere = strWhere & " AND ( "
	
	For pvCnt = 0 To 9
	    tmpArrSQL(pvCnt) = " ITEM_CD_" & CStr(pvCnt + 1) & " = " & FilterVar(strItemCd, "''", "S")
	Next 	
		        
	strWhere = strWhere & Join(tmpArrSQL, " Or ") & ")"
	
	If 	CommonQueryRs2by2(strSelect, " Y_CAST (NOLOCK) ", strWhere, lgF0) = False Then
		Call DisplayMsgBox("Y60040","X", Frm1.vspdData.Text,"X")
		Call LookUpCastCdFail(Frm1.vspdData.Text, Row)	    
		Exit Function
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	With frm1.vspdData	
	
		.Col = C_CastNm
		.text = lgF0(1)
		.Col = C_Cavi
		.text = lgF0(3)
		strCaviQty = lgF0(3)
		.Col = C_CurCount
		.text = lgF0(2)
		
		Call .GetText(C_ProdQty, Row, strProdQty)
		
		If Not (Trim(strCaviQty) = "0") Then
			strInputQty =  Cint(Abs(strProdQty) / uniCDBl(strCaviQty) + uniCDBl(0.499999))
		Else
			strInputQty = "0"	
		End If	
		
		Call .SetText(C_InputQty, Row, strInputQty)
			
	End With
	
	Call LookUpCastCdSuccess(Row)

End Function

Function LookUpCastCdFail(Byval strCastCd, Byval Row)


    With frm1.vspdData
		.Row = Row
		.Col = C_CastCd
		.text = ""
		.Col = C_CastNm
		.text = ""
		.Col = C_Cavi
		.text = "0"
		.Col = C_CurCount
		.text = "0"
	
	End With
	
	Call SetActiveCell(frm1.vspdData, C_CastCd, Row, "M","X","X")
	Set gActiveElement = document.activeElement
End Function

Function LookUpCastCdSuccess(Byval Row)
	Call SetActiveCell(frm1.vspdData, C_CastCd, Row, "M","X","X")
	Set gActiveElement = document.activeElement
	
End Function


'-------------------------------------  LookUpTopCastCd()  -----------------------------------------
'	Name : LookUpTopCastCd()
'	Description : LookUp Item By Plant
'--------------------------------------------------------------------------------------------------------- 
Function LookUpTopCastCd(Byval Row)
    
	Dim strVal
	Dim strSelect, strWhere
	Dim lgF0
	Dim pvCnt
	Dim tmpArrSQL(9)
	Dim strDelFlag, strItemCd, strWcCd
	Dim gComNum1000, gComNumDec, gAPNum1000, gAPNumDec
	
	Dim strInputQty, strProdQty, strCaviQty
	
	gComNum1000 = parent.gComNum1000
	gComNumDec = parent.gComNumDec
	gAPNum1000 = parent.gAPNum1000
	gAPNumDec = parent.gAPNumDec
	
	Call frm1.vspdData.GetText(C_DelFlag, Row, strDelFlag)
	
	If Ucase(Trim(strDelFlag)) = "Y" Then Exit Function
	
	Call frm1.vspdData.GetText(C_ItemCd, Row, strItemCd)
	Call frm1.vspdData.GetText(C_WcCd, Row, strWcCd)
	
	strSelect = " TOP 1 CAST_NM, CUR_ACCNT, PRS_UNIT, CAST_CD  "
	
	strWhere = " USE_YN = 'Y' "
	strWhere = strWhere & " AND SET_PLANT = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
	strWhere = strWhere & " AND SET_PLACE = " & FilterVar(strWcCd, "''", "S")
	strWhere = strWhere & " AND ( "
	
	For pvCnt = 0 To 9
	    tmpArrSQL(pvCnt) = " ITEM_CD_" & CStr(pvCnt + 1) & " = " & FilterVar(strItemCd, "''", "S")
	Next 	
		        
	strWhere = strWhere & Join(tmpArrSQL, " Or ") & ")"
	
	If 	CommonQueryRs2by2(strSelect, " Y_CAST (NOLOCK) ", strWhere & " ORDER BY CAST_CD ", lgF0) = False Then
		Exit Function
	End If
	
	lgF0 = Split(lgF0, Chr(11))
	
	With frm1.vspdData	
		
		.Col = C_CastCd
		.text = UCase(Trim(lgF0(4)))
		.Col = C_CastNm
		.text = lgF0(1)
		.Col = C_Cavi
		.text = lgF0(3)
		strCaviQty = lgF0(3)
		.Col = C_CurCount
		.text = lgF0(2)
		
		Call .GetText(C_ProdQty, Row, strProdQty)
		
		If Not (Trim(strCaviQty) = "0") Then
			strInputQty =  Cint(Abs(strProdQty) / uniCDBl(strCaviQty) + uniCDBl(0.499999))
		Else
			strInputQty = "0"	
		End If	
		
		Call .SetText(C_InputQty, Row, strInputQty)
			
	End With
	
	
End Function



'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenOprRef()">��������</A> </TD>
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
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������"></OBJECT>');</SCRIPT>
									</TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtFromItemCd.value, 1 ">&nbsp;<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=25 tag="14" ALT="ǰ���">&nbsp;~</TD>						
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProdOrderNo() "></TD>
								</TR>								
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtToItemCd.value , 2">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=25 tag="14" ALT="ǰ���"></TD>
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprCd" SIZE=8 MAXLENGTH=3 tag="11xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprCd()">
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value,0"></TD>								
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConWC()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="14" ALT="�۾����"></TD>
									<TD CLASS=TD5 NOWRAP>������������</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboCastRsltFlg" ALT="������������" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
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
								<TD HEIGHT="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=5 WIDTH=100% colspan=4></TD>
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
			  <TD WIDTH="*" align="left"><!--<a><button name="btnAutoSel" class="clsmbtn">��ü����</button></a>--></TD>
			  <TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
      </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hToItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hOprCD" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hResultFlg" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>