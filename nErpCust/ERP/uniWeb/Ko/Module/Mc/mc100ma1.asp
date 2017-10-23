<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC100MA1
'*  4. Program Name         : �������ô���� 
'*  5. Program Desc         : �������ô���� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003-04-08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ahn Jung Je
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
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   =====================================-->

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
Const BIZ_PGM_ID = "mc100mb1.asp"											
Const BIZ_PGM_JUMP_ID = "M3111MA1"
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

Dim C_Select		
Dim C_ProdOrderNo	'����������ȣ 
Dim C_ItemCd		'ǰ�� 
Dim C_ItemNm		'ǰ��� 
Dim C_Spec			'�԰� 
Dim C_ReqDt			'�ʿ��� 
Dim C_Unit			'������ 
Dim C_ReqQty		'�ʿ���� 
Dim C_PoBaseQty		'���ּ��� 
Dim C_PoNeedQty		'�����ʿ���� 
Dim C_PoUnit		'���ִ��� 
Dim C_PoUnitQty		'���ִ������ּ��� 
Dim C_TrackingNo	'Tracking No
Dim C_WCCd			'�۾���  
Dim C_PlanStartDt	'���������� 
Dim C_PlanComptDt	'�ϷΌ���� 
Dim C_OprNo			'���� 
Dim C_Seq			'��ǰ������� 
Dim C_SubSeq		'�������ü��� 
Dim C_ReleaseDt

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop    
Dim strDate
Dim iDBSYSDate

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE  
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0          
    lgStrPrevKey = ""          
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtFromReqDt.text	= UniConvDateAToB(UNIDateAdd ("D", -7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtToReqDt.text	= UniConvDateAToB(UNIDateAdd ("D", 7, LocSvrDate, parent.gServerDateFormat), parent.gServerDateFormat, parent.gDateFormat)
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
		ggoSpread.Spreadinit "V20030226", ,Parent.gAllowDragDropSpread
				
		.ReDraw = false
				
		.MaxCols = C_ReleaseDt + 1    
		.MaxRows = 0    
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetCheck C_Select, "����",8,,,true
		ggoSpread.SSSetEdit 	C_ProdOrderNo,  "����������ȣ"	,20
		ggoSpread.SSSetEdit 	C_ItemCd,       "ǰ��"			,20
		ggoSpread.SSSetEdit 	C_ItemNm,       "ǰ���"		,25
		ggoSpread.SSSetEdit 	C_Spec,			"�԰�"			,25
		ggoSpread.SSSetEdit 	C_ReqDt,		"�ʿ���"		,10, 2
		ggoSpread.SSSetEdit 	C_Unit,			"������"		,10
		ggoSpread.SSSetFloat	C_ReqQty,		"�ʿ䷮"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_PoBaseQty,	"���ּ���",15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_PoNeedQty,	"�����ʿ����"	,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_PoUnit,		"���ִ���"		,10
		ggoSpread.SSSetFloat	C_PoUnitQty,	"���ִ������ּ���"	,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No"	,25
		ggoSpread.SSSetEdit 	C_WcCd,			"�۾���"		,20
		ggoSpread.SSSetEdit 	C_PlanStartDt,  "����������"	,11, 2
		ggoSpread.SSSetEdit 	C_PlanComptDt,	"�ϷΌ����"	,11, 2
		ggoSpread.SSSetEdit 	C_OprNo,		"����"			,12
		ggoSpread.SSSetEdit 	C_Seq,			"��ǰ�������"	,12
		ggoSpread.SSSetEdit 	C_SubSeq,		"�������ü���"	,12
		ggoSpread.SSSetEdit 	C_ReleaseDt,	"�۾�������"	,11, 2
		
		Call ggoSpread.SSSetColHidden(C_Seq, .MaxCols,	True)
		Call SetSpreadLock 
		
		.ReDraw = true    
    End With
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.SpreadLock	-1, -1
    ggoSpread.SpreadUnLock	C_Select, -1, C_Select, -1
End Sub

'========================== 2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2110", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDoTime, lgF0, lgF1, Chr(11))
    
    frm1.cboDoTime.value = ""
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
	C_Unit			= 7
	C_ReqQty		= 8
	C_PoBaseQty		= 9
	C_PoNeedQty		= 10
	C_PoUnit		= 11
	C_PoUnitQty		= 12
	C_TrackingNo	= 13
	C_WCCd			= 14
	C_PlanStartDt	= 15
	C_PlanComptDt	= 16
	C_OprNo			= 17
	C_Seq			= 18
	C_SubSeq		= 19
	C_ReleaseDt		= 20
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
			C_Unit			= iCurColumnPos(7)
			C_ReqQty		= iCurColumnPos(8)
			C_PoBaseQty		= iCurColumnPos(9)
			C_PoNeedQty		= iCurColumnPos(10)
			C_PoUnit		= iCurColumnPos(11)
			C_PoUnitQty		= iCurColumnPos(12)
			C_TrackingNo	= iCurColumnPos(13)
			C_WCCd			= iCurColumnPos(14)
			C_PlanStartDt	= iCurColumnPos(15)
			C_PlanComptDt	= iCurColumnPos(16)
			C_OprNo			= iCurColumnPos(17)
			C_Seq			= iCurColumnPos(18)
			C_SubSeq		= iCurColumnPos(19)
			C_ReleaseDt		= iCurColumnPos(20)
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
Function OpenProdOrderNo(i)
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If i = 1 then		
		If IsOpenPop = True Or UCase(frm1.txtProdOrderNo1.className) = "PROTECTED" Then Exit Function
	Else
		If IsOpenPop = True Or UCase(frm1.txtProdOrderNo2.className) = "PROTECTED" Then Exit Function
	End if
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	Else
		If Plant_Check() = False Then Exit Function
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
	arrParam(3) = "RL"									'From Status
	arrParam(4) = "ST"									'To Status
	If i = 1 then	
		arrParam(5) = Trim(frm1.txtProdOrderNo1.value)
	Else
		arrParam(5) = Trim(frm1.txtProdOrderNo2.value)
	End if
	
	arrParam(6) = ""		
	arrParam(7) = ""		
	arrParam(8) = ""			
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If i = 1 then	
			frm1.txtProdOrderNo1.focus
		Else
			frm1.txtProdOrderNo2.focus
		End if	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		If i = 1 then	
			frm1.txtProdOrderNo1.Value	= arrRet(0)
			frm1.txtProdOrderNo1.focus
		Else
			frm1.txtProdOrderNo2.Value	= arrRet(0)
			frm1.txtProdOrderNo2.focus
		End if	
	End If
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	Else
		If Plant_Check() = False Then Exit Function
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
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	frm1.vspdData.Row =	frm1.vspdData.ActiveRow 
	

		
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
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

'===========================================================================
' Function Name : JumpChgCheck
' Function Desc : Jump�� ����Ÿ ���濩�� üũ 
'=========================================================================== 
Function JumpChgCheck()
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	Call PgmJump(BIZ_PGM_JUMP_ID)
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
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables

    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtFromReqDt.focus 
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
End Sub

'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Select
			frm1.vspdData.Row = i
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_Select, i, 1)
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Select
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_Select, i, 0)
		Next	
	End If
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
'   Event Name : txtFromReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromReqDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToReqDt_KeyDown(keycode, shift)
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
    If CheckRunningBizProcess = True Then Exit Sub							'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        
    If OldLeft <> NewLeft Then Exit Sub
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then									'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
End Sub

<%
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
%>
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Select And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
	    Case 0

			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row 
			frm1.vspdData.text = "" 
			lgBlnFlgChgValue = False					
	    End Select
	End If
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

    If frm1.vspdData.MaxRows = 0 Then Exit Sub                                                    'If there is no data.

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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False														'��: Processing is NG
    Err.Clear																'��: Protect system from crashing
	
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If

    If Not chkfield(Document, "1") Then	Exit Function									'��: This function check indispensable field
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables		
    
	If Trim(frm1.txtItemCd.value) <> "" Then
		If Plant_Item_Check() = False Then Exit Function
	Else
		frm1.txtItemNm.value = ""
		If Plant_Check() = False Then Exit Function
	End If
	
	If ValidDateCheck(frm1.txtFromReqDt, frm1.txtToReqDt) = False Then Exit Function
												'��: Initializes local global variables
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

'===========================================  5.1.2 FncNew()  ===========================================
'=	Event Name : FncNew																					=
'=	Event Desc : This function is related to New Button of Main ToolBar									=
'========================================================================================================
Function FncNew()
	Dim IntRetCD 

	FncNew = False								

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then Exit Function
	End If

	Call ggoOper.ClearField(Document, "A")			
	Call ggoOper.LockField(Document, "N")			
	Call SetDefaultVal
	Call SetToolBar("11100000000011")				

	Call InitVariables
		
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If								

	Set gActiveElement = document.activeElement
	FncNew = True									
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear         
    
    If CheckRunningBizProcess = True Then Exit Function
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then                 
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")   
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                        
    If Not ggoSpread.SSDefaultCheck Then Exit Function         
    
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
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
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
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
    
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")           
		
		If IntRetCD = vbNo Then Exit Function
		
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
			strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtFromReqDt="		& Trim(.hFromReqDt.value)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToReqDt="		& Trim(.hToReqDt.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.hItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtProdOrderNo1="	& UCase(Trim(.hProdOrderNo1.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdOrderNo2="	& UCase(Trim(.hProdOrderNo2.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey							'
			strVal = strVal & "&lgIntFlgMode="		& lgIntFlgMode
			strVal = strVal & "&txtMaxRows="		& .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd="		& UCase(Trim(.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtFromReqDt="		& Trim(.txtFromReqDt.text)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtToReqDt="		& Trim(.txtToReqDt.text)			'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd="			& UCase(Trim(.txtItemCd.value))			'��: ��ȸ ���� ����Ÿ		
			strVal = strVal & "&txtProdOrderNo1="	& UCase(Trim(.txtProdOrderNo1.value))	'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdOrderNo2="	& UCase(Trim(.txtProdOrderNo2.value))	'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey
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
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data Save 
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim strVal
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

    DbSave = False    

    igColSep = Parent.gColSep
    igRowSep = Parent.gRowSep
	frm1.txtMode.value = Parent.UID_M0002
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
   
    If LayerShowHide(1) = False Then Exit Function
	
	With frm1
		For lRow = 1 To .vspdData.MaxRows
		
			If Trim(GetSpreadText(.vspdData,0,lRow,"X","X")) = ggoSpread.UpdateFlag then
				strVal = "U" & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ProdOrderNo,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X"))    & igColSep
				strVal = strVal & UNICDbl(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X"))    & igColSep
				strVal = strVal & UNICDbl(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_TrackingNo,lRow,"X","X"))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_ReqDt,lRow,"X","X")))    & igColSep
				strVal = strVal & UNICDbl(GetSpreadText(frm1.vspdData,C_ReqQty,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Unit,lRow,"X","X"))    & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_WCCd,lRow,"X","X"))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_PlanStartDt,lRow,"X","X")))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_PlanComptDt,lRow,"X","X")))    & igColSep
				strVal = strVal & UNIConvDate(Trim(GetSpreadText(frm1.vspdData,C_ReleaseDt,lRow,"X","X")))    & igColSep
				strVal = strVal & lRow & igRowSep

				lGrpCnt = lGrpCnt + 1
			End If
			
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
	End With
	
	frm1.txtMaxRows.value = lGrpCnt - 1
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                       
End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()										
	Call InitVariables()
	Call MainQuery()
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

'========================================================================================
' Function Name : Plant_Check
' Function Desc : 
'========================================================================================
Function Plant_Check()
	Plant_Check = False

	'-----------------------
	'Check Plant CODE		'�����ڵ尡 �ִ� �� üũ 
	'-----------------------
    If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus 
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	
	Plant_Check = True
End Function

'========================================================================================
' Function Name : Plant_Item_Check
' Function Desc : 
'========================================================================================
Function Plant_Item_Check()
	Plant_Item_Check = False

	'-----------------------
	'Check Item CODE		'�����ڵ尡 �ִ� �� üũ 
	'-----------------------
    If 	CommonQueryRs(" C.PLANT_NM, B.ITEM_NM "," B_ITEM_BY_PLANT A, B_ITEM B, B_PLANT C ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD " & _
						" AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			frm1.txtPlantCd.focus 
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtPlantNm.Value = lgF0(0)
	
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)
			Call DisplayMsgBox("122700","X","X","X")
			frm1.txtItemCd.focus 
		Else
			Call DisplayMsgBox("122600","X","X","X")
			frm1.txtItemNm.Value = ""
			frm1.txtItemCd.focus 
		End If
		
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
    lgF1 = Split(lgF1, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	frm1.txtItemNm.Value = lgF1(0)
	
	Plant_Item_Check = True
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������ô�� ����</font></td>
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
									<TD CLASS=TD5 NOWRAP>�ʿ���</TD> 
									<TD CLASS="TD6">
										<script language =javascript src='./js/mc100ma1_OBJECT1_txtFromReqDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/mc100ma1_OBJECT2_txtToReqDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td NOWRAP>
													<INPUT TYPE=TEXT NAME="txtProdOrderNo1" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo(1)"></TD>
												<td NOWRAP> &nbsp;~ &nbsp;</td>
											</tr>
										</table>
									</TD>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MaxLength=40 tag="14"></TD>
									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo2" SIZE=20 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo(2)"></TD>
									
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
				<TD WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
						    <script language =javascript src='./js/mc100ma1_OBJECT3_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
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
					    <button name="btnSelect" class="clsmbtn">�ϰ�����</button>&nbsp;
					    <BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">�ϰ��������</BUTTON>
					</td>
					<TD WIDTH="*" ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">���ֵ��</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromReqDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hToReqDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo1" tag="24">
<INPUT TYPE=HIDDEN NAME="hProdOrderNo2" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

