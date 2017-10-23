<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc800qb1
'*  4. Program Name         : ����������Ȳ��ȸ 
'*  5. Program Desc         : ����������Ȳ��ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/24
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Woo Guen
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
<!-- '#########################################################################################################
'												1. �� �� �� 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   =====================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
Const BIZ_PGM_QRY_ID = "mc800qb1.asp"								'��: Head Query �����Ͻ� ���� ASP�� 

Dim C_ProdtOrderNo		
Dim C_ItemCd					
Dim C_ItemNm					
Dim C_Specification	
Dim C_ReqDt	
Dim C_ReqQty		
Dim C_BaseUnit	
Dim C_DoQty		 
Dim C_RcptQty		 
Dim C_BpCd		 
Dim C_BpNm
Dim C_DoDt
Dim C_DoTime	
Dim C_DoTimeDesc	
Dim C_DoStatus	
Dim C_DoStatusDesc
Dim C_TrackingNo				
Dim C_PoNo				
Dim C_PoSeqNo				
Dim C_DoQtyPoUnit			
Dim C_RcptQtyPoUnit		
Dim C_PoUnit		
Dim C_OprNo		
Dim C_Seq					
Dim C_SubSeq						
Dim C_WcCd						
Dim C_WcNm	
Dim C_PlanStartDt			
Dim C_PlanComptDt				
Dim C_ReleaseDt	

'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop										'Popup

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False					'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2115", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboDlvyOrderStatus, lgF0, lgF1, Chr(11))
End Sub

'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim LocSvrDate
	
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtReqFromDt.text = UniConvDateAToB(UNIDateAdd ("D", -5, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtReqToDt.text   = UniConvDateAToB(UNIDateAdd ("D", 10, LocSvrDate, Parent.gServerDateFormat), Parent.gServerDateFormat, parent.gDateFormat)
	Call SetToolbar("1100000000001111")
End Sub
   
'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    
    Call InitSpreadPosVariables()
    
    With frm1
    
	    ggoSpread.Source = .vspdData
	    ggoSpread.Spreadinit "V20030107", , Parent.gAllowDragDropSpread

	 	.vspdData.ReDraw = false
	    .vspdData.MaxCols = C_ReleaseDt + 1
	    .vspdData.MaxRows = 0

	    Call GetSpreadColumnPos("A")

	    ggoSpread.SSSetEdit		C_ProdtOrderNo, "����������ȣ", 16,,,,2
	    ggoSpread.SSSetEdit		C_ItemCd,		"ǰ��", 20,,,,2
	    ggoSpread.SSSetEdit		C_ItemNm,		"ǰ���", 25
	    ggoSpread.SSSetEdit		C_Specification,"�԰�", 25
	    ggoSpread.SSSetDate 	C_ReqDt,		"�ʿ���", 12, 2, parent.gDateFormat    
		ggoSpread.SSSetFloat	C_ReqQty,		"�ʿ����",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_BaseUnit,		"�ʿ����", 10
	    ggoSpread.SSSetFloat	C_DoQty,		"�ʿ䳳�����ü���",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_RcptQty,		"�ʿ�����԰����",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_BpCd,			"����ó", 10
	    ggoSpread.SSSetEdit		C_BpNm,			"����ó��", 20
	    ggoSpread.SSSetDate 	C_DoDt,			"����������", 12, 2, parent.gDateFormat    
		ggoSpread.SSSetCombo	C_DoTime,		"�������ýð�", 04
		ggoSpread.SSSetEdit		C_DoTimeDesc,	"�������ýð�", 12
	    ggoSpread.SSSetCombo	C_DoStatus,		"�������û���", 04
		ggoSpread.SSSetEdit		C_DoStatusDesc, "�������û���", 12
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
		ggoSpread.SSSetEdit 	C_PoNo,			"���ֹ�ȣ", 20
		ggoSpread.SSSetEdit 	C_PoSeqNo,		"���ּ���", 10,1
		ggoSpread.SSSetFloat	C_DoQtyPoUnit,	"���ֳ������ü���",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetFloat	C_RcptQtyPoUnit,"���ִ����԰����",15,Parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    ggoSpread.SSSetEdit		C_PoUnit,		"���ִ���", 10
	    ggoSpread.SSSetEdit		C_OprNo,		"����", 10
	    ggoSpread.SSSetEdit		C_Seq,			"��ǰ�������", 10
	    ggoSpread.SSSetEdit		C_SubSeq,		"�������ü���", 10
	    ggoSpread.SSSetEdit		C_WcCd,			"�۾���", 10
	    ggoSpread.SSSetEdit		C_WcNm,			"�۾����", 16
	    ggoSpread.SSSetDate 	C_PlanStartDt,	"������ȹ����", 10, 2, parent.gDateFormat
	    ggoSpread.SSSetDate 	C_PlanComptDt,	"�Ϸ��ȹ����", 10, 2, parent.gDateFormat
	    ggoSpread.SSSetDate 	C_ReleaseDt,	"�۾�������", 10, 2, parent.gDateFormat
    
	    Call ggoSpread.SSSetColHidden(C_DoTime, C_DoTime, True)
	    Call ggoSpread.SSSetColHidden(C_DoStatus, C_DoStatus, True)
	    Call ggoSpread.SSSetColHidden(C_Seq, C_Seq, True)
	    Call ggoSpread.SSSetColHidden(C_SubSeq, C_SubSeq, True)
	    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		
		.vspdData.ReDraw = true
		
		  ggoSpread.Source = frm1.vspdData
		  ggoSpread.SpreadLockWithOddEvenRowColor()

    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables()
	C_ProdtOrderNo					= 1
	C_ItemCd						= 2
	C_ItemNm						= 3
	C_Specification					= 4
	C_ReqDt							= 5	
	C_ReqQty						= 6
	C_BaseUnit						= 7
	C_DoQty							= 8
	C_RcptQty						= 9 
	C_BpCd							= 10
	C_BpNm							= 11
	C_DoDt							= 12
	C_DoTime						= 13
	C_DoTimeDesc					= 14
	C_DoStatus						= 15
	C_DoStatusDesc					= 16
	C_TrackingNo					= 17	
	C_PoNo							= 18
	C_PoSeqNo						= 19
	C_DoQtyPoUnit					= 20
	C_RcptQtyPoUnit					= 21
	C_PoUnit						= 22
	C_OprNo							= 23
	C_Seq							= 24
	C_SubSeq						= 25		
	C_WcCd							= 26	
	C_WcNm							= 27
	C_PlanStartDt					= 28
	C_PlanComptDt					= 29	
	C_ReleaseDt						= 30
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
			
			C_ProdtOrderNo				= iCurColumnPos(1)
			C_ItemCd					= iCurColumnPos(2)
			C_ItemNm					= iCurColumnPos(3)
			C_Specification				= iCurColumnPos(4)  
			C_ReqDt						= iCurColumnPos(5)  
			C_ReqQty					= iCurColumnPos(6)  
			C_BaseUnit					= iCurColumnPos(7)  
			C_DoQty						= iCurColumnPos(8)  
			C_RcptQty					= iCurColumnPos(9)  
			C_BpCd						= iCurColumnPos(10) 
			C_BpNm						= iCurColumnPos(11) 
			C_DoDt						= iCurColumnPos(12) 
			C_DoTime					= iCurColumnPos(13) 
			C_DoTimeDesc				= iCurColumnPos(14) 
			C_DoStatus					= iCurColumnPos(15) 
			C_DoStatusDesc				= iCurColumnPos(16) 
			C_TrackingNo				= iCurColumnPos(17) 
			C_PoNo						= iCurColumnPos(18) 
			C_PoSeqNo					= iCurColumnPos(19) 
			C_DoQtyPoUnit				= iCurColumnPos(20) 
			C_RcptQtyPoUnit				= iCurColumnPos(21) 
			C_PoUnit					= iCurColumnPos(22) 
			C_OprNo						= iCurColumnPos(23) 
			C_Seq						= iCurColumnPos(24) 
			C_SubSeq					= iCurColumnPos(25) 
			C_WcCd						= iCurColumnPos(26) 
			C_WcNm						= iCurColumnPos(27) 
			C_PlanStartDt				= iCurColumnPos(28) 
			C_PlanComptDt				= iCurColumnPos(29) 
			C_ReleaseDt					= iCurColumnPos(30)
 	End Select
End Sub
 
'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
	arrField(0) = "PLANT_CD"					' Field��(0)
	arrField(1) = "PLANT_NM"					' Field��(1)
	
	arrHeader(0) = "����"					' Header��(0)
	arrHeader(1) = "�����"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
		frm1.txtPlantCd.focus    	
		Set gActiveElement = document.activeElement	
	End If
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim iCalledAspName
	Dim arrParam(5), arrField(2)
	
	Dim IntRetCD
	Dim arrRet

	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    '���������� �ʿ��մϴ� 
		frm1.txtPlantCd.focus
		Exit Function
	End If

    '���� üũ �Լ� ȣ�� 
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	'------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 
	arrField(2) = 3 ' -- Spec

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.value = arrRet(0)
		frm1.txtItemNm.value = arrRet(1) 
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function

'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Biz Partner PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtBpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "����ó"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "����ó"				
    arrHeader(1) = "����ó��"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus		
		Exit Function
	Else
		frm1.txtBpCd.Value    = arrRet(0)		
		frm1.txtBpNm.Value    = arrRet(1)	
		frm1.txtBpCd.focus		
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNo()
	Dim iCalledAspName
	Dim arrParam(8)
	
	Dim arrRet
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtReqFromDt.Text
	arrParam(2) = frm1.txtReqToDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "RL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value) 
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = ""
	'arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = ""
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtProdOrderNo.focus
		Exit Function
	Else
		frm1.txtProdOrderNo.Value    = arrRet(0)
		frm1.txtProdOrderNo.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim iCalledAspName
	Dim arrParam(2)
	
	Dim strRet
	Dim IntRetCD

	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	arrParam(0) = "N"	'Return Flag
	arrParam(1) = "N"	'Release Flag
	arrParam(2) = ""	'STO Flag

	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus		
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus		
		Set gActiveElement = document.activeElement
	End If	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim arrParam(4)
	
	Dim arrRet
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If IsOpenPop = True  Then Exit Function
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = frm1.txtReqFromDt.Text
	arrParam(4) = frm1.txtReqToDt.Text	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
		Set gActiveElement = document.activeElement	
	End If
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     				'��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          			'��: Lock  Suitable  Field
    Call InitSpreadSheet 

	Call SetDefaultVal
	Call InitVariables		'��: Initializes local global variables
 	Call InitComboBox

	 'Plant Code, Plant Name Setting 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
         
    With frm1.vspdData
		If .MaxRows <= 0 Then Exit Sub
		If Row < 1 Then
			ggoSpread.Source = frm1.vspdData
			 
 			If lgSortKey = 1 Then
 				ggoSpread.SSSort Col					'Sort in Ascending
 				lgSortKey = 2
 			Else
 				ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 				lgSortKey = 1
			End If 
		End If
	
    End With
	
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then Exit Sub			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
    If OldLeft <> NewLeft Then Exit Sub
     '----------  Coding part  -------------------------------------------------------------
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
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
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtReqFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReqFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqFromDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtReqFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtReqToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReqToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtReqToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishStartDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� MainQuery�Ѵ�.
'=======================================================================================================
Sub txtReqFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtFinishEndDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� MainQuery�Ѵ�.
'=======================================================================================================
Sub txtReqToDt_KeyDown(keycode, shift)
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
   
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	If frm1.txtBpCd.value = "" Then
		frm1.txtBpNm.value = "" 
	End If
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then Exit Function										'��: This function check indispensable field

    If ValidDateCheck(frm1.txtReqFromDt, frm1.txtReqToDt) = False Then Exit Function
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables															'��: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function										'��: Query db data
       
    Set gActiveElement = document.ActiveElement   
    FncQuery = True																'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()                                           '��: Protect system from crashing
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)												'��: ȭ�� ���� 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
   
    Err.Clear							'��: Protect system from crashing

    DbQuery = False                                                         			'��: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
 
    Dim strVal
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'��: 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'��: ���� ���� ����Ÿ 
		strVal = strVal & "&txtReqFromDt=" & Trim(frm1.hReqFromDt.value)		'��: ���� ���� ����Ÿ 
		strVal = strVal & "&txtReqToDt=" & Trim(frm1.hReqToDt.value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)				'��: ���� ���� ����Ÿ 
		strVal = strVal & "&txtBpCd=" & Trim(frm1.hBpCd.value)					'��: ���� ���� ����Ÿ		
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.hProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtPoNo=" & Trim(frm1.hPoNo.value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)		'��: ���� ���� ����Ÿ  
		strVal = strVal & "&cboDlvyOrderStatus=" & Trim(frm1.hDlvyOrderStatus.value)	'��: ���� ���� ����Ÿ		  
	Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & Parent.UID_M0001						'��: 
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ���� ���� ����Ÿ 
		strVal = strVal & "&txtReqFromDt=" & Trim(frm1.txtReqFromDt.text)		'��: ���� ���� ����Ÿ 
		strVal = strVal & "&txtReqToDt=" & Trim(frm1.txtReqToDt.text)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)				'��: ���� ���� ����Ÿ 
		strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.value)					'��: ���� ���� ����Ÿ		
		strVal = strVal & "&txtProdOrderNo=" & Trim(frm1.txtProdOrderNo.value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)					'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)		'��: ���� ���� ����Ÿ  
		strVal = strVal & "&cboDlvyOrderStatus=" & Trim(frm1.cboDlvyOrderStatus.value)	'��: ���� ���� ����Ÿ	  
	End If    

    Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                          	'��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(LngMaxRow)													'��: ��ȸ ������ ������� 
	
	Call SetToolBar("11000000000111")											'��: ��ư ���� ���� 
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field

	If frm1.vspdData.MaxRows <= 0 Then Exit Function

    lgIntFlgMode = Parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
	
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
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
'######################################################################################################### -->
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����������Ȳ��ȸ</font></td>
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
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlantCd()">
														 <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>�ʿ���</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/mc800qa1_OBJECT1_txtReqFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/mc800qa1_OBJECT2_txtReqToDt.js'></script>
									</TD>
								</TR>
								<TR><TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd">
														 <INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����ó" NAME="txtBpCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�������� ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="�������� ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo"  SIZE=20 MAXLENGTH=18 ALT="���ֹ�ȣ" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
								</TR>
								<TR><TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo frm1.txtTrackingNo.value,0"></TD>
									<TD CLASS="TD5" NOWRAP>�������û���</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDlvyOrderStatus" ALT="�������û���" STYLE="Width: 165px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
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
								<TD HEIGHT=* WIDTH=100%>
									<script language =javascript src='./js/mc800qa1_vspdData_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMajorFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hReqToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24"><INPUT TYPE=HIDDEN NAME="hDlvyOrderStatus" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
