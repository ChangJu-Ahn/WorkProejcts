<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC400MA1
'*  4. Program Name         : ����������� 
'*  5. Program Desc         : ����������� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003-02-25
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ryu Sung Won
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
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_QRY_ID	= "mc400mb1.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "mc400mb2.asp"							'��: �����Ͻ� ���� ASP�� 

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

' Grid 1(vspdData) - Operation
Dim C_Select		
Dim C_ProdOrderNo	'����������ȣ 
Dim C_ItemCd		'ǰ�� 
Dim C_ItemNm		'ǰ��� 
Dim C_Spec			'�԰� 
Dim C_ReqDt			'�ʿ��� 
Dim C_ReqQty		'�ʿ䷮ 
Dim C_BaseUnit		'�ʿ���� 
Dim C_DoQty			'�ʿ䳳�����÷� 
Dim C_BpCd			'����ó 
Dim C_BpNm			'����ó�� 
Dim C_PoNo			'���ֹ�ȣ 
Dim C_PoSeqNo		'���ּ��� 
Dim C_DoQtyPoUnit	'���ֳ������÷� 
Dim C_PoUnit		'���ִ��� 
Dim C_OprNo			'���� 
Dim C_Seq			'��ǰ������� 
Dim C_SubSeq		'�������ü��� 
Dim C_WcCd			'�۾��� 
Dim C_WcNm			'�۾���� 
Dim C_PlanStartDt	'���������� 
Dim C_PlanComptDt	'�ϷΌ���� 
Dim C_ReleaseDt		'�۾������� 
Dim C_TrackingNo	'TrackingNo

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							<%'Variable is for Dirty flag%>
Dim lgIntGrpCount								<%'Group View Size�� ������ ���� %>
Dim lgIntFlgMode								<%'Variable is for Operation Status%>
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgSortKey

Dim strDate
Dim iDBSYSDate

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgStrPrevKey4 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgSortKey = 1
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtDoDate.text		= UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)
	
	Call SetToolbar("1110000000001111")
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","M","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021122", ,Parent.gAllowDragDropSpread
				
		.ReDraw = false
				
		.MaxCols = C_TrackingNo + 1    
		.MaxRows = 0    
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck	C_Select ,		"����"			,8,,,1
		ggoSpread.SSSetEdit 	C_ProdOrderNo,  "����������ȣ"	,20
		ggoSpread.SSSetEdit 	C_ItemCd,       "ǰ��"			,20
		ggoSpread.SSSetEdit 	C_ItemNm,       "ǰ���"		,20
		ggoSpread.SSSetEdit 	C_Spec,			"�԰�"			,20
		ggoSpread.SSSetEdit 	C_ReqDt,		"�ʿ���"		,12, 2
		ggoSpread.SSSetFloat	C_ReqQty,		"�ʿ䷮"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BaseUnit,		"�ʿ����"		,8
		ggoSpread.SSSetFloat	C_DoQty,		"�������÷�",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_BpCd,			"����ó"		,15
		ggoSpread.SSSetEdit 	C_BpNm,			"����ó��"		,20
		ggoSpread.SSSetEdit 	C_PoNo,			"���ֹ�ȣ"		,20
		ggoSpread.SSSetEdit 	C_PoSeqNo,		"���ּ���"		,10, 1
		ggoSpread.SSSetFloat	C_DoQtyPoUnit,	"���ִ����������÷�",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_PoUnit,		"���ִ���"		,8
		ggoSpread.SSSetEdit 	C_OprNo,		"����"			,8
		ggoSpread.SSSetEdit 	C_Seq,			"��ǰ�������"	,16
		ggoSpread.SSSetEdit 	C_SubSeq,		"�������ü���"	,16
		ggoSpread.SSSetEdit 	C_WcCd,			"�۾���"		,10
		ggoSpread.SSSetEdit 	C_WcNm,			"�۾����"		,20
		ggoSpread.SSSetEdit 	C_PlanStartDt,  "����������"	,12, 2
		ggoSpread.SSSetEdit 	C_PlanComptDt,	"�ϷΌ����"	,12, 2
		ggoSpread.SSSetEdit 	C_ReleaseDt,	"�۾�������"	,12, 2
		ggoSpread.SSSetEdit 	C_TrackingNo,	"Tracking No."	,20
		
		Call ggoSpread.SSSetColHidden(C_Seq,		C_Seq,		True)
		Call ggoSpread.SSSetColHidden(C_SubSeq,		C_SubSeq,	True)
		Call ggoSpread.SSSetColHidden(.MaxCols,		.MaxCols,	True)
		
		ggoSpread.SSSetSplit2(3)
		
		Call SetSpreadLock 
		
		.ReDraw = true    
    End With
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	Dim i
	ggoSpread.Source = frm1.vspdData
	
	For i=2 To frm1.vspdData.MaxCols
		ggoSpread.SSSetProtected i, -1, -1
	Next
End Sub

'========================== 2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2110", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDoTime, lgF0, lgF1, Chr(11))
End Sub

'============================  2.2.7 InitSpreadPosVariables() ===========================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()
	C_Select		= 1
	C_ProdOrderNo	= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5
	C_ReqDt			= 6
	C_ReqQty		= 7
	C_BaseUnit		= 8
	C_DoQty			= 9
	C_BpCd			= 10
	C_BpNm			= 11
	C_PoNo			= 12
	C_PoSeqNo		= 13
	C_DoQtyPoUnit	= 14
	C_PoUnit		= 15
	C_OprNo			= 16
	C_Seq			= 17
	C_SubSeq		= 18
	C_WcCd			= 19
	C_WcNm			= 20
	C_PlanStartDt	= 21
	C_PlanComptDt	= 22
	C_ReleaseDt		= 23
	C_TrackingNo	= 24
End Sub

'============================  2.2.8 GetSpreadColumnPos()  ==============================
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_Select		= iCurColumnPos(1)
			C_ProdOrderNo	= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5)
			C_ReqDt			= iCurColumnPos(6)
			C_ReqQty		= iCurColumnPos(7)
			C_BaseUnit		= iCurColumnPos(8)
			C_DoQty			= iCurColumnPos(9)
			C_BpCd			= iCurColumnPos(10)
			C_BpNm			= iCurColumnPos(11)
			C_PoNo			= iCurColumnPos(12)
			C_PoSeqNo		= iCurColumnPos(13)
			C_DoQtyPoUnit	= iCurColumnPos(14)
			C_PoUnit		= iCurColumnPos(15)
			C_OprNo			= iCurColumnPos(16)
			C_Seq			= iCurColumnPos(17)
			C_SubSeq		= iCurColumnPos(18)
			C_WcCd			= iCurColumnPos(19)
			C_WcNm			= iCurColumnPos(20)
			C_PlanStartDt	= iCurColumnPos(21)
			C_PlanComptDt	= iCurColumnPos(22)
			C_ReleaseDt		= iCurColumnPos(23)
			C_TrackingNo	= iCurColumnPos(24)
    End Select
End Sub    

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
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
	End If	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function

	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""									'ProdFromDt
	arrParam(2) = ""									'ProdToDt
	arrParam(3) = ""									'From Status
	arrParam(4) = ""									'To Status
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = ""	
	arrParam(8) = ""									'cboOrderType
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtProdOrderNo.focus
		Exit Function
	Else
	    frm1.txtProdOrderNo.Value	= arrRet(0)
		frm1.txtProdOrderNo.focus
	End If
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	 If Trim(frm1.txtPlantCd.value) = "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	'iCalledAspName = AskPRAspName("B1B11PA1")
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
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
	End If	
End Function

'--------------------------------------  OpenTrackingInfo()  ---------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingInfo()
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = "PROTECTED" Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = ""	
	arrParam(4) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
	End If
End Function

'--------------------------------------  OpenPoNo()  -----------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
		Dim strRet
		Dim arrParam(2)
		Dim iCalledAspName
		Dim IntRetCD
		
		If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
		IsOpenPop = True
		
		arrParam(0) = "N"	'Return Flag
		arrParam(1) = "Y"	'Release Flag
		arrParam(2) = ""	'STO Flag
		
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
			frm1.txtPoNo.focus
			Exit Function
		Else
			frm1.txtPoNo.value = strRet(0)
			frm1.txtPoNo.focus
		End If	
		
End Function

'--------------------------------------  OpenBpInfo()  -----------------------------------------------------
'	Name : OpenBpInfo()
'	Description : Open Bp Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBpInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtBpCd.Value)
	arrParam(3) = ""	'Trim(frm1.txtBpNm.Value)
	
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
	End If	
End Function

'------------------------------------------  OpenWcCd()  -------------------------------------------------
'	Name : OpenWcCd()
'	Description : Work Center Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWcCd()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	 If Trim(frm1.txtPlantCd.value) = "" Then
		Call parent.DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�۾����˾�"	
	arrParam(1) = "P_WORK_CENTER"				
	arrParam(2) = frm1.txtWcCd.value  
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & parent.FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
				 
	arrParam(5) = "�۾���"			
	
    arrField(0) = "WC_CD"	
    arrField(1) = "WC_NM"	
    arrField(2) = "INSIDE_FLG"
    arrField(3) = "WC_MGR"	
    
    arrHeader(0) = "�۾���"		
    arrHeader(1) = "�۾����"		
    arrHeader(2) = "�۾���Ÿ��"		
    arrHeader(3) = "�۾�������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtWcCd.focus
		Exit Function
	Else
		frm1.txtWcCd.value = arrRet(0)
		frm1.txtWcNm.value = arrRet(1)
		frm1.txtWcCd.focus
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    Call InitComboBox

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.cboDoTime.focus 
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement

End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFromReqDt.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtFromReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToReqDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtToReqDt.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtToReqDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtDoDate_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDoDate_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDoDate.Action = 7 
	    Call SetFocusToDocument("M")  
        frm1.txtDoDate.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtFromReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtDoDate_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If CheckRunningBizProcess = True Then							'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then									'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
    Else
        
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	lgBlnFlgChgValue = True    
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData
		.Row = Row
		.Col = C_Select
		
		ggoSpread.Source = frm1.vspdData
		
		If .Text = "Y" Then
			If ButtonDown = 0 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If
		Else
			If ButtonDown = 1 Then
				ggoSpread.UpdateRow Row
			Else
				ggoSpread.SSDeleteFlag Row,Row
			End If			
		End If
	End With
End Sub

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    'On Error Resume Next                                   
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                  
    Call ggoOper.ClearField(Document, "2")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.cboDoTime.focus 
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
    FncNew = True                                                           

End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False														'��: Processing is NG
    Err.Clear																'��: Protect system from crashing

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	If Trim(frm1.txtPlantCd.value) = "" Then frm1.txtPlantNm.value = "" 
	If Trim(frm1.txtItemCd.value) = "" Then frm1.txtItemNm.value = "" 	
	If Trim(frm1.txtBpCd.value) = "" Then frm1.txtBpNm.value = "" 	
	If Trim(frm1.txtWcCd.value) = "" Then frm1.txtWcNm.value = "" 
	
    '-----------------------
    'Erase contents area
    '-----------------------
    'Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables														'��: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then										'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'��: Query db data
	End If
	
	Set gActiveElement = document.activeElement
    FncQuery = True															'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    On Error Resume Next                                                    '��: Protect system from crashing
    Err.Clear
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '��: No data changed!!
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Trim(frm1.txtPlantCd.value) = "" Then frm1.txtPlantNm.value = "" 
	If Trim(frm1.txtItemCd.value) = "" Then frm1.txtItemNm.value = "" 	
	If Trim(frm1.txtBpCd.value) = "" Then frm1.txtBpNm.value = "" 	
	If Trim(frm1.txtWcCd.value) = "" Then frm1.txtWcNm.value = "" 
	    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then   
		Exit Function           
    End If     							                                      '��: Save db data
    
	Set gActiveElement = document.activeElement
    FncSave = True 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'��: Protect system from crashing
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
	Set gActiveElement = document.activeElement
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

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
    Dim strYear1
    Dim strMonth1
    Dim strDay1
    Dim strDate1
       
    DbQuery = False
    
	Call LayerShowHide(1)
    
	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDoDate="			& Trim(.hDoDate.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&cboDoTime="			& Trim(.hDoTime.value)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.hItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtBpCd="			& UCase(Trim(.hBpCd.value))				'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.hProdOrderNo.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPoNo="			& UCase(Trim(.hPoNo.value))				'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtWcCd="			& UCase(Trim(.hWcCd.value))				'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtTrackingNo="		& UCase(Trim(.hTrackingNo.value))		'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&lgStrPrevKey1="		& lgStrPrevKey1							'
			strVal = strVal & "&lgStrPrevKey2="		& lgStrPrevKey2							'
			strVal = strVal & "&lgStrPrevKey3="		& lgStrPrevKey3
			strVal = strVal & "&lgStrPrevKey4="		& lgStrPrevKey4							'
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtDoDate="			& Trim(.txtDoDate.text)					'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&cboDoTime="			& Trim(.cboDoTime.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.txtItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtBpCd="			& UCase(Trim(.txtBpCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtProdOrderNo="	& UCase(Trim(.txtProdOrderNo.value))	'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtPoNo="			& UCase(Trim(.txtPoNo.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtWcCd="			& UCase(Trim(.txtWcCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtTrackingNo="		& UCase(Trim(.txtTrackingNo.value))		'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&lgStrPrevKey1="		& ""
			strVal = strVal & "&lgStrPrevKey2="		& ""
			strVal = strVal & "&lgStrPrevKey3="		& ""
			strVal = strVal & "&lgStrPrevKey4="		& ""
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		End If
	End With

    Call RunMyBizASP(MyBizASP, strVal)														'��: �����Ͻ� ASP �� ���� 
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("11101000000111")														'��: ��ư ���� ���� 
	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True	
    frm1.vspdData.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
	On Error Resume Next
	Err.Clear

    Dim lRow 
    Dim strVal
	Dim lGrpCnt
	Dim igColSep,igRowSep
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
	
    DbSave = False                                                          '��: Processing is NG
	
    LayerShowHide(1)
	
	igColSep = parent.gColSep
	igRowSep = parent.gRowSep
	strVal = ""
	lGrpCnt = 1
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferMaxCount = -1 
	iTmpDBufferMaxCount = -1 
	    
	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

    With frm1
		.txtMode.Value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.Value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
		
    For lRow = 1 To .vspdData.MaxRows
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
            Case ggoSpread.UpdateFlag
				
				strVal = "U" & igColSep & lRow & igColSep			'��: U=Update
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProdOrderNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X"))    & igRowSep
				
                lGrpCnt = lGrpCnt + 1
        End Select
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.UpdateFlag
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
    Next

	.txtMaxRows.value = lGrpCnt-1
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: ���� �����Ͻ� ASP �� ���� 
    DbSave = True                                                           '��: Processing is OK
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	frm1.txtPlantCd.value = frm1.hPlantCd.value
	
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
	IsOpenPop = False
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
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
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    ggoSpread.Source = gActiveSpdSheet
		
	Call ggoSpread.ReOrderingSpreadData()
End Sub 

'------------------------------------------  ChkBtnAll()  --------------------------------------------------
'	Name : ChkBtnAll()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Sub ChkBtnAll()
	Dim LngRow
	If frm1.vspdData.MaxRows <= 0 Then Exit Sub
	
	With frm1.vspdData
		For LngRow = 1 To .MaxRows
			.Row = LngRow
			.Col = C_Select
			.Value = 1

			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow LngRow
		Next
	End With
	lgBlnFlgChgValue = True
End Sub


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������Ȯ�����</font></td>
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
			 						<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>����������</TD> 
									<TD CLASS=TD6 NOWRAP>
										<TABLE><TR>
										<TD>
										<script language =javascript src='./js/mc400ma1_OBJECT3_txtDoDate.js'></script>
										</TD>
										<TD>
										&nbsp;
										<SELECT NAME="cboDoTime" ALT="�������ýð�" STYLE="Width: 98px;" tag="12"></SELECT>
										</TD>
										</TR></TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MaxLength=40 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=15 MAXLENGTH=10 tag="11xxxU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MaxLength=40 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�������� ��ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="�������� ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>
									<TD CLASS=TD5 NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="���ֹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=10 MAXLENGTH=20 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWcCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 MaxLength=40 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=20 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" colspan=4>
									<script language =javascript src='./js/mc400ma1_OBJECT4_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnCopy" CLASS="CLSMBTN" Flag=1 ONCLICK=ChkBtnAll>��ü����</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hDoDate" tag="24">
<INPUT TYPE=HIDDEN NAME="hDoTime" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hPoNo" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

