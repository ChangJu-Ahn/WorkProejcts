<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : mc700ma1
'*  4. Program Name         : �������ø��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/03/13
'*  8. Modified date(Last)  : 2003/03/13
'*  9. Modifier (First)     : Kang Su Hwan
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*							  
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
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
Const BIZ_PGM_ID = "mc700mb1.asp"											
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Dim C_Check
Dim C_ProdtOrderNo		'����������ȣ 
Dim C_ItemCd			'ǰ�� 
Dim C_ItemNm			'ǰ��� 
Dim C_ItemSpec			'�԰� 
Dim C_ReqDt				'�ʿ��� 
Dim C_DoQty				'�������÷� 
Dim C_RcptQty			'�԰���� 
Dim C_BaseUnit			'������ 
Dim C_SupplierCd		'���޾�ü 
Dim C_SupplierNm		'���޾�ü�� 
Dim C_TrackingNo		'TrackingNo
Dim C_OprNo				'���� 
Dim C_PoNo				'���ֹ�ȣ 
Dim C_PoSeqNo			'���ּ��� 
Dim C_Seq				'��ǰ�����Ϸù�ȣ 
Dim C_SubSeq			'�������ü��� 
Dim C_DoStatusDesc		'�������û��� 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim StartDate,EndDate
 
EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("d", -7, EndDate, Parent.gServerDateFormat)
EndDate = UNIDateAdd("d", +7, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop          
'==========================================   Selection()  ======================================
'	Name : Selection()
'	Description : �ϰ����ù�ư�� Event �ռ� 
'================================================================================================
Sub Selection()
	Dim index,Count
 
 	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Check
		
		if frm1.vspdData.Text = "1" then
			frm1.vspdData.Text = "0"
		else
			frm1.vspdData.Text = "1"
		End if
		
		frm1.vspdData.Col = 0 
		
		if ggoSpread.UpdateFlag = frm1.vspdData.Text then    
			frm1.vspdData.Text=""
	    else
	    	ggoSpread.UpdateRow Index
		End if
	Next 
	
	frm1.vspdData.ReDraw = true
End Sub


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

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Set gActiveElement = document.activeElement
	frm1.txtPlantCd.value = Parent.gPlant
	frm1.txtPlantNm.value = Parent.gPlantNm
	
    frm1.txtFrDt.Text = StartDate
    frm1.txtToDt.Text = EndDate
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True

	Call SetToolbar("1110000000001111")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_Check			= 1
	C_ProdtOrderNo	= 2			'����������ȣ 
	C_ItemCd		= 3			'ǰ�� 
	C_ItemNm		= 4			'ǰ��� 
	C_ItemSpec		= 5			'�԰� 
	C_ReqDt			= 6
	C_DoQty			= 7			'�������÷� 
	C_RcptQty		= 8			'�԰���� 
	C_BaseUnit		= 9 		'������ 
	C_SupplierCd	= 10		'���޾�ü 
	C_SupplierNm	= 11		'���޾�ü�� 
	C_TrackingNo	= 12		'TrackingNo
	C_OprNo			= 13		'���� 
	C_PoNo			= 14		'���ֹ�ȣ 
	C_PoSeqNo		= 15		'���ּ��� 
	C_Seq			= 16		'��ǰ�����Ϸù�ȣ 
	C_SubSeq		= 17		'�������ü��� 
	C_DoStatusDesc  = 18
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData

    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030303",,Parent.gAllowDragDropSpread  

	.ReDraw = false	
	
    .MaxCols = C_DoStatusDesc + 1														
	.Col = .MaxCols:    .ColHidden = True											
    .MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetCheck C_Check, "����",10,,,true
	ggoSpread.SSSetEdit C_ProdtOrderNo,"����������ȣ",18
	ggoSpread.SSSetEdit C_ItemCd, "ǰ��",18
	ggoSpread.SSSetEdit C_ItemNm, "ǰ���",20
	ggoSpread.SSSetEdit C_ItemSpec, "ǰ��԰�",20
    ggoSpread.SSSetDate C_ReqDt	,"�ʿ���", 10, 2, Parent.gDateFormat
    SetSpreadFloatLocal	C_DoQty, "�������÷�", 15, 1, 3
    SetSpreadFloatLocal	C_RcptQty, "�԰����", 15, 1, 3
    ggoSpread.SSSetEdit C_BaseUnit, "����", 6
    ggoSpread.SSSetEdit C_SupplierCd, "����ó", 10
    ggoSpread.SSSetEdit C_SupplierNm, "����ó��", 20
    ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 25
	ggoSpread.SSSetEdit C_OprNo,"����",7
	ggoSpread.SSSetEdit C_PoNo,"���ֹ�ȣ",18
    ggoSpread.SSSetEdit C_PoSeqNo, "���ּ���", 10
    ggoSpread.SSSetEdit C_Seq, "��ǰ�����Ϸù�ȣ", 10
    ggoSpread.SSSetEdit C_SubSeq, "�������� ����", 10
    ggoSpread.SSSetEdit	C_DoStatusDesc, "�������û���", 12
    
	Call ggoSpread.SSSetColHidden(C_Seq,		C_Seq,		True)
	Call ggoSpread.SSSetColHidden(C_SubSeq,		C_SubSeq,		True)

	.ReDraw = true
	
    Call SetSpreadLock 
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock	-1, -1
    ggoSpread.SpreadUnLock	C_Check, -1, C_Check, -1
    
    .vspdData.ReDraw = True

    End With
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

			C_Check			= iCurColumnPos(1)
			C_ProdtOrderNo	= iCurColumnPos(2)			'����������ȣ 
			C_ItemCd		= iCurColumnPos(3)			'ǰ�� 
			C_ItemNm		= iCurColumnPos(4)			'ǰ��� 
			C_ItemSpec		= iCurColumnPos(5)			'�԰� 
			C_ReqDt			= iCurColumnPos(6)			'�ʿ��� 
			C_DoQty			= iCurColumnPos(7)			'�������÷� 
			C_RcptQty		= iCurColumnPos(8)			'�԰���� 
			C_BaseUnit		= iCurColumnPos(9)		'������ 
			C_SupplierCd	= iCurColumnPos(10)		'���޾�ü 
			C_SupplierNm	= iCurColumnPos(11)		'���޾�ü�� 
			C_TrackingNo	= iCurColumnPos(12)		'TrackingNo
			C_OprNo			= iCurColumnPos(13)		'���� 
			C_PoNo			= iCurColumnPos(14)		'���ֹ�ȣ 
			C_PoSeqNo		= iCurColumnPos(15)		'���ּ��� 
			C_Seq			= iCurColumnPos(16)		'��ǰ�����Ϸù�ȣ 
			C_SubSeq		= iCurColumnPos(17)		'�������ü��� 
			C_DoStatusDesc  = iCurColumnPos(18)		'�������û��� 
	End Select

End Sub	

'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"						
	arrParam(1) = "B_PLANT"      					
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		
	arrParam(4) = ""								
	arrParam(5) = "����"						
	
    arrField(0) = "PLANT_CD"						
    arrField(1) = "PLANT_NM"						
    
    arrHeader(0) = "����"						
    arrHeader(1) = "�����"						
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim iCalledAspName,IntRetCD
	Dim arrParam(5), arrField(2)

	If IsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
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
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then	
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenSupplier()  -------------------------------------------------
'	Name : OpenSupplier()
'	Description : OpenSupplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"						
	arrParam(1) = "B_BIZ_PARTNER"					

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
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
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : ���Ÿ� ���� �׸����� ���� �κ��� ����ȸ� �� �Լ��� ���� �ؾ���.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              'Lot ���� Maker Lot ���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6"				  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999"
    End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                         
    Call ggoOper.LockField(Document, "N")       
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                        
    Call InitVariables                          
    Call SetDefaultVal
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtSupplierCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
End Sub

'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : ���ó�� ��ư�� Ŭ���� ��� �߻� 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Check
			frm1.vspdData.Row = i
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_Check, i, 1)
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : ���ó����� ��ư�� Ŭ���� ��� �߻� 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_Check
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_Check, i, 0)
		Next	
	End If
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
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
<%
'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'==========================================================================================
%>
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Check And Row > 0 Then
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

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	Call SetPopupMenuItemInf("0000111111")
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    		
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
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

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

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
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                               
	
	ggoSpread.Source = frm1.vspdData
	
    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
   '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then						
       Exit Function
    End If
    
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","������", "X")			
			Exit Function
		End if   
	End with
	
	If Not chkConditions Then
		Exit Function
	End If
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    Set gActiveElement = document.ActiveElement   
    FncQuery = True											
    
End Function

'========================================================================================
' Function Name : chkConditions
' Function Desc : 
'========================================================================================
Function chkConditions()
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	chkConditions=False

	If Len(Trim(frm1.txtPlantCd.Value))  Then
		Call CommonQueryRs(" PLANT_NM ", " B_PLANT ", " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iNameArr = Split(lgF0, Chr(11))
		If Err.number <> 0 Then
			MsgBox Err.Description, vbInformation, parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		If lgF0="" Then 
			Call DisplayMsgBox("970000", "X","����", "X")	
			Exit Function
		End If
		frm1.txtPlantNm.Value = iNameArr(0)
	End If
	If Len(Trim(frm1.txtSupplierCd.Value)) Then
		Call CommonQueryRs(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD =  " & FilterVar(frm1.txtSupplierCd.Value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iNameArr = Split(lgF0, Chr(11))
		If Err.number <> 0 Then
			MsgBox Err.Description, vbInformation, parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		If lgF0="" Then 
			Call DisplayMsgBox("970000", "X","����ó", "X")	
			Exit Function
		End If
		frm1.txtSupplierNM.Value = iNameArr(0)
	End If
	If Len(Trim(frm1.txtItemCd.Value)) Then
		Call CommonQueryRs(" ITEM_NM ", " B_ITEM ", " ITEM_CD =  " & FilterVar(frm1.txtItemCd.Value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		iNameArr = Split(lgF0, Chr(11))
		If Err.number <> 0 Then
			MsgBox Err.Description, vbInformation, parent.gLogoName
			Err.Clear 
			Exit Function
		End If
		If lgF0="" Then 
			Call DisplayMsgBox("970000", "X","ǰ��", "X")	
			Exit Function
		End If
		frm1.txtItemNm.Value = iNameArr(0)
	End If

	chkConditions=True	
End Function

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
		frm1.txtSupplierCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If
    FncNew = True                                                           

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       
    
    Err.Clear                                                               
    'On Error Resume Next                                                   
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                         
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    If DbDelete = False Then Exit Function
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  
    Call ggoOper.ClearField(Document, "2")                                  
    
    Set gActiveElement = document.ActiveElement   
    FncDelete = True                                                        
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear    
    
    If CheckRunningBizProcess = True Then
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
    If DbSave = False Then Exit Function
    
    Set gActiveElement = document.ActiveElement   
    FncSave = True                                     
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                 
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
	if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_MULTI)						
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)                
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")              
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Set gActiveElement = document.ActiveElement   
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
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    Err.Clear                                         

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode = " & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd	=" & .hdnPlantCd.value 
	    strVal = strVal & "&txtSupplier	=" & .hdnSupplier.value 
		strVal = strVal & "&txtFrDt		=" & .hdnFrDt.value
		strVal = strVal & "&txtToDt		=" & .hdnToDt.value
	    strVal = strVal & "&txtItemCd	=" & .hdnItemCd.value 
		strVal = strVal & "&txtMaxRows	=" & frm1.vspdData.MaxRows
	else
	    strVal = BIZ_PGM_ID & "?txtMode	=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd	=" & .txtPlantCd.value 
	    strVal = strVal & "&txtSupplier	=" & .txtSupplierCd.value 
		strVal = strVal & "&txtFrDt		=" & .txtFrDt.Text
		strVal = strVal & "&txtToDt		=" & .txtToDt.Text
	    strVal = strVal & "&txtItemCd	=" & .txtItemCd.value 
		strVal = strVal & "&txtMaxRows	=" & frm1.vspdData.MaxRows
	end if 
	
	Call RunMyBizASP(MyBizASP, strVal)				
        
    End With
    
    DbQuery = True
End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()									
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE							
    
    Call ggoOper.LockField(Document, "Q")				
	Call SetSpreadLock
	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False
	Call SetToolBar("11101000000111")														'��: ��ư ���� ���� 
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim igColSep, igRowSep
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
    
    If LayerShowHide(1) = False Then Exit Function
    
	With frm1
		.txtMode.value = Parent.UID_M0002
    
		igColSep = parent.gColSep
		igRowSep = parent.gRowSep

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
		lGrpCnt = 1
		strVal = ""

		For lRow = 1 To .vspdData.MaxRows
		    If Trim(GetSpreadText(frm1.vspdData,C_Check,lRow,"X","X")) = "1" Then
				strVal = Trim(GetSpreadText(frm1.vspdData,C_ProdtOrderNo,lRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OprNo,lRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_Seq,lRow,"X","X")) & igColSep
				strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SubSeq,lRow,"X","X")) & igColSep				
				strVal = strVal & lRow & igRowSep
				lGrpCnt = lGrpCnt + 1
			End If
			
			Select Case Trim(GetSpreadText(frm1.vspdData,C_Check,lRow,"X","X"))
			    Case "1"
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

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	
	End With
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�������ø���</font></td>
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
					    <TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
											   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>����ó</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó"  NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
           									   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
					</TR>					   
					<TR>
						<TD CLASS=TD5 NOWRAP>�ʿ���</TD>
						<TD CLASS=TD6 NOWRAP>								
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/mc700ma1_fpDateTime1_txtFrDt.js'></script>
									</td>
									<td NOWRAP>&nbsp;~&nbsp;</td>
									<td NOWRAP>
										<script language =javascript src='./js/mc700ma1_fpDateTime1_txtToDt.js'></script>
									</td>
								<tr>
							</table>
						<TD CLASS="TD5" NOWRAP>ǰ��</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="ǰ��" NAME="txtItemCd" SIZE=10 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">
											   <INPUT TYPE=TEXT Alt="ǰ��" NAME="txtItemNm" SIZE=20 tag="14"></TD>
					</TR>
					</TABLE></FIELDSET></TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
						    <script language =javascript src='./js/mc700ma1_I150130379_vspdData.js'></script>
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
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">

</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
