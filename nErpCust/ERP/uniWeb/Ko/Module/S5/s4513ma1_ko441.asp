<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************2007-12-26**************************
'*  1. Module Name          : ����
'*  2. Function Name        : 
'*  3. Program ID           : S4513MA1_KO441
'*  4. Program Name         : �������Ͻ�����ȸ
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/12/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<html>

<head>
<title><%=Request("strASPMnuMnuNm")%></title>
<!-- '******************************************  1.1 Inc ����   **********************************************
' ���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<link rel="stylesheet" type="Text/css" href="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<script language="VBScript" src="../../inc/incCliMAMain.vbs"></script>
<script language="VBScript" src="../../inc/incCliMAEvent.vbs"></script>
<script language="VBScript" src="../../inc/incCliVariables.vbs"></script>
<script language="VBScript" src="../../inc/incCliMAOperation.vbs"></script>
<script language="VBScript" src="../../inc/incCliRdsQuery.vbs"></script>
<script language="VBScript" src="../../inc/incHRQuery.vbs"></script>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<script language="javascript" src="../../inc/TabScript.js"> </script>
<script language="VBScript">

Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->

<!--'******************************************  1.2 Global ����/��� ����  ***********************************
' 1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************!-->
Const BIZ_PGM_ID = "s4513mb1_ko441.asp"

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================!-->

Dim C_BP_NM				'CUST
Dim C_ITEM_GP			'ǰ��׷�
Dim C_Total				'�� �Ѽ���
Dim C_Qty_01			'���ϼ���(1��)
Dim C_Qty_02			'���ϼ���(2��)
Dim C_Qty_03			'���ϼ���(3��)
Dim C_Qty_04			'���ϼ���(4��)
Dim C_Qty_05			'���ϼ���(5��)
Dim C_Qty_06			'���ϼ���(6��)
Dim C_Qty_07			'���ϼ���(7��)
Dim C_Qty_08			'���ϼ���(8��)
Dim C_Qty_09			'���ϼ���(9��)
Dim C_Qty_10			'���ϼ���(10��)
Dim C_Qty_11			'���ϼ���(11��)
Dim C_Qty_12			'���ϼ���(12��)
Dim C_Qty_13			'���ϼ���(13��)
Dim C_Qty_14			'���ϼ���(14��)
Dim C_Qty_15			'���ϼ���(15��)
Dim C_Qty_16			'���ϼ���(16��)
Dim C_Qty_17			'���ϼ���(17��)
Dim C_Qty_18			'���ϼ���(18��)
Dim C_Qty_19			'���ϼ���(19��)
Dim C_Qty_20			'���ϼ���(20��)
Dim C_Qty_21			'���ϼ���(21��)
Dim C_Qty_22			'���ϼ���(22��)
Dim C_Qty_23			'���ϼ���(23��)
Dim C_Qty_24			'���ϼ���(24��)
Dim C_Qty_25			'���ϼ���(25��)
Dim C_Qty_26			'���ϼ���(26��)
Dim C_Qty_27			'���ϼ���(27��)
Dim C_Qty_28			'���ϼ���(28��)
Dim C_Qty_29			'���ϼ���(29��)
Dim C_Qty_30			'���ϼ���(30��)
Dim C_Qty_31			'���ϼ���(31��)

Dim C_MaxCol
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================  
Dim IsOpenPop  
Dim lgIsOpenPop   
'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================  
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                    ' �������( create )
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0          
    lgStrPrevKey = ""     
    lgStrPrevKeyIndex = ""     
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
End Sub

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
 
Sub initSpreadPosVariables()

	C_BP_NM		= 1			
	C_ITEM_GP	= 2		
	C_Total		= 3			
	C_Qty_01	= 4		
	C_Qty_02	= 5		
	C_Qty_03	= 6		
	C_Qty_04	= 7		
	C_Qty_05	= 8		
	C_Qty_06	= 9		
	C_Qty_07	= 10	
	C_Qty_08	= 11
	C_Qty_09	= 12
	C_Qty_10	= 13
	C_Qty_11	= 14
	C_Qty_12	= 15
	C_Qty_13	= 16
	C_Qty_14	= 17
	C_Qty_15	= 18
	C_Qty_16	= 19
	C_Qty_17	= 20		
	C_Qty_18	= 21		
	C_Qty_19	= 22		
	C_Qty_20	= 23		
	C_Qty_21	= 24		
	C_Qty_22	= 25		
	C_Qty_23	= 26		
	C_Qty_24	= 27		
	C_Qty_25	= 28		
	C_Qty_26	= 29		
	C_Qty_27	= 30		
	C_Qty_28	= 31		
	C_Qty_29	= 32		
	C_Qty_30	= 33		
	C_Qty_31	= 34		
	  
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= --!>

 Sub SetDefaultVal()
	
	Call SetToolBar("1100000000101111")

	frm1.txtDocumentDt.Text = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat) 
 	Call ggoOper.FormatDate(frm1.txtDocumentDt, parent.gDateFormat, 2)
	
	frm1.txtPlantCd1.value=parent.gPlant
	frm1.txtPlantNm1.value=parent.gPlantNm 
    frm1.txtPlantCd1.focus 
	frm1.cboApType.value = "1"
	C_MaxCol	= 34	
	Set gActiveElement = document.activeElement

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================

Sub InitSpreadSheet()

	Call initSpreadPosVariables()    

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

	With frm1.vspdData
		
        .ReDraw = False
        '.MaxCols = C_QTY_31 + 1		
		.MaxCols = C_MaxCol + 1	
        .MaxRows = 0 		
       ' Call GetSpreadColumnPos()
		
		ggoSpread.SSSetEdit			C_BP_NM		,		"CUST"			, 20
		ggoSpread.SSSetEdit			C_ITEM_GP	,		"����"			, 20
		ggoSpread.SSSetFloat		C_Total		,		"�� ����"		, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_01	,		"1��"			, 11 , Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_02	,		"2��"			, 11 , Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_03	,		"3��"			, 11 , Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_04	,		"4��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_05	,		"5��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_06	,		"6��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_07	,		"7��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_08	,		"8��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_09	,		"9��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_10	,		"10��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_11	,		"11��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_12	,		"12��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_13	,		"13��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_14	,		"14��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_15	,		"15��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_16	,		"16��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_17	,		"17��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_18	,		"18��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_19	,		"19��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_20	,		"20��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_21	,		"21��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_22	,		"22��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_23	,		"23��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_24	,		"24��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_25	,		"25��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_26	,		"26��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_27	,		"27��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_28	,		"28��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_29	,		"29��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_30	,		"30��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Qty_31	,		"31��"			, 11, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		
		
		Call ggoSpread.SSSetSplit2(3)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)		

        Call SetSpreadLock
        
		.ReDraw = true

	End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos
	
   
    ggoSpread.Source = frm1.vspdData
	
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
	
	 C_BP_NM		=	iCurColumnPos(1) 
	 C_ITEM_GP		=	iCurColumnPos(2) 
	 C_Total		=	iCurColumnPos(3) 
	 C_Qty_01		=	iCurColumnPos(4) 
	 C_Qty_02		=	iCurColumnPos(5) 
	 C_Qty_03		=	iCurColumnPos(6) 
	 C_Qty_04		=	iCurColumnPos(7) 
	 C_Qty_05		=	iCurColumnPos(8) 
	 C_Qty_06		=	iCurColumnPos(9) 
	 C_Qty_07		=	iCurColumnPos(10) 
	 C_Qty_08		=	iCurColumnPos(11) 
	 C_Qty_09		=	iCurColumnPos(12) 
	 C_Qty_10		=	iCurColumnPos(13) 
	 C_Qty_11		=	iCurColumnPos(14) 
	 C_Qty_12		=	iCurColumnPos(15) 
	 C_Qty_13		=	iCurColumnPos(16) 
	 C_Qty_14		=	iCurColumnPos(17) 
	 C_Qty_15		=	iCurColumnPos(18) 
	 C_Qty_16		=	iCurColumnPos(19) 
	 C_Qty_17		=	iCurColumnPos(20) 
	 C_Qty_18		=	iCurColumnPos(21) 
	 C_Qty_19		=	iCurColumnPos(22) 
	 C_Qty_20		=	iCurColumnPos(23) 
	 C_Qty_21		=	iCurColumnPos(24) 
	 C_Qty_22		=	iCurColumnPos(25) 
	 C_Qty_23		=	iCurColumnPos(26) 
	 C_Qty_24		=	iCurColumnPos(27) 
	 C_Qty_25		=	iCurColumnPos(28) 
	 C_Qty_26		=	iCurColumnPos(29) 
	 C_Qty_27		=	iCurColumnPos(30) 
	 C_Qty_28		=	iCurColumnPos(31) 
	 C_Qty_29		=	iCurColumnPos(32) 
	 C_Qty_30		=	iCurColumnPos(33) 
	 C_Qty_31		=	iCurColumnPos(34) 
			 
	
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
    
        ggoSpread.SpreadLockWithOddEvenRowColor()
        						         
        .vspdData.ReDraw = True
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

Sub SetSpreadColor1(ByVal pvStarRow, Byval pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With frm1
    	.vspdData.ReDraw = False

    	.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
       Next
    End If   
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	arrParam(0) = "����"
	arrParam(1) = "B_Plant"

	arrParam(2) = Trim(frm1.txtPlantCd1.Value)

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
		frm1.txtPlantCd1.Value= arrRet(0)
		frm1.txtPlantNm1.Value= arrRet(1)
		frm1.txtPlantCd1.focus
	End If
End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================  
 Sub Form_Load()

    Call LoadInfTB19029                  
    Call ggoOper.LockField(Document, "N")   
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart) 	 
	Call SetDefaultVal
    Call InitSpreadSheet      
    Call InitVariables
	Call InitComboBox
End Sub


'==========================================================================================
'   Event Name : txtDocumentDt    
'   Event Desc :
'==========================================================================================

 Sub txtDocumentDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDocumentDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtDocumentDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================

Sub txtDocumentDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

'	Call SetPopupMenuItemInf("1101111111")

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
		Exit Sub
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub InitComboBox()
	With frm1
		Call SetCombo(frm1.cboApType, "0","0����")
		Call SetCombo(frm1.cboApType, "1","1����") 
		Call SetCombo(frm1.cboApType, "2","2����")      
		Call SetCombo(frm1.cboApType, "3","3����")
		Call SetCombo(frm1.cboApType, "4","4����") 
		Call SetCombo(frm1.cboApType, "5","5����")      
	     .cboApType.value = "1"
    End With
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

Function OpenConPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)
	
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1
		Select Case pvIntWhere
			Case C_PopPlant		'���� 
				iArrParam(1) = "dbo.B_PLANT"									
				iArrParam(2) = Trim(.txtConPlant.value)				
				iArrParam(3) = ""										
				iArrParam(4) = ""										
				
				iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"
							
				iArrHeader(0) = .txtPlant1.alt						
				iArrHeader(1) = .txtPlantNm1.alt					
	
				.txtPlant1.focus			
			
		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' �˾� Title
	iArrParam(5) = iArrHeader(0)							' ��ȸ���� ��Ī 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopUp = SetConPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function SetConPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant
				.txtPlant1.value = pvArrRet(0)
				.txtPlantNm1.value = pvArrRet(1) 		
				
		End Select
	End With

End Function
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 

	Dim IntRetCD 
    Err.Clear                                               
    
    FncQuery = False        
                                        
    '-----------------------
    'Check previous data area
    '-----------------------
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO  , "X", "X")
		
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
	    
	With frm1
		IF Trim(.txtDocumentDt.text) = "" Then
			Call DisplayMsgBox("17a003", "X","���", "X")
			Exit Function
		End if 
		
		Dim Yr
		Dim Mnth
		
		Yr = Left(.txtDocumentDt.DateValue,4)
		Mnth = Mid(.txtDocumentDt.DateValue,5, 2)
	
		If Mnth = "01" or Mnth = "03" or Mnth = "05" or Mnth = "07" or Mnth = "08" or Mnth = "10" or Mnth = "12" Then
			C_MaxCol = C_Qty_31 '32
		ElseIf Mnth = "02" Then
			If CInt(Yr) Mod 4 = 0 Then				'������ ��� 2���� 29�Ϸ� ó�� 				
				C_MaxCol = C_Qty_29  '30
			Else
				C_MaxCol = C_Qty_28  '29
			End If
		Else
			C_MaxCol = C_Qty_30 '31
		End If		
	
		.vspdData.focus
		ggoSpread.Source = .vspdData		
		
	End with
	
	Call InitSpreadSheet
	
	Call ggoOper.LockField(Document, "Q")        
	Call SetToolBar("1100000000011111")	
	
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False Then 
	   Exit Function
	END IF
	
	FncQuery = True  
	Set gActiveElement = document.activeElement

End Function



'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData.Maxrows < 1 then exit function
	frm1.vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                               
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Curr,C_Cost,"C" ,"I","X","X")
	Set gActiveElement = document.activeElement
	frm1.vspdData.ReDraw = True
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
 
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(parent.C_SINGLEMULTI)   
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(parent.C_SINGLEMULTI , False)      
	Set gActiveElement = document.activeElement
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
 	Dim IntRetCD

	FncExit = False
	 
	ggoSpread.Source = frm1.vspdData      
	 
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")      
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	Set gActiveElement = document.activeElement
	FncExit = True    
End Function

'===================================  BtnPreview()  ========================================
Function BtnPreview() 
	
    If Not chkField(Document, "1") Then	
       Exit Function
    End If
	
    'IF ChkKeyField() = False Then 
	'	Exit Function
    'End if

	dim var1,var2,var3,var4,var5
	
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim ObjName

	'with frm1
	'	if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
	'		Call DisplayMsgBox("17a003", "X","������", "X")	
	'		Exit Function
	'	End if   
	'End with				

	var1 = FilterVar(Trim(UCase(frm1.txtPlantCd1.value)), "''" ,  "S") 
	strUrl = strUrl & "plant_cd|" & var1 

	
	ObjName = AskEBDocumentName("mz190ma1_ko100","ebr")
	Call FncEBRPreview(ObjName, strUrl)
	Call BtnDisabled(0)

	Set gActiveElement = document.activeElement
End Function


'===================================  FncBtnPrint()  ========================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt

	dim var1,var2,var3,var4,var5
    Dim ObjName
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If

    IF ChkKeyField() = False Then 
		Exit Function
    End if

	On Error Resume Next                  
	
	lngPos = 0

	with frm1
		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","������", "X")	
			Exit Function
		End if   
	End with		
		
	var1 = FilterVar(Trim(UCase(frm1.txtPlantCd1.value)), "''" ,  "S") 
	
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
	
	strUrl = strUrl & "plant_cd|" & var1 

	ObjName = AskEBDocumentName("mz190ma1_ko100","ebr")

	Call FncEBRprint(EBAction, ObjName, strUrl)
	
	Call BtnDisabled(0)
		
	Set gActiveElement = document.activeElement
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	ChkKeyField = true
	
	strWhere = " PLANT_CD = '" & FilterVar(frm1.txtPlantCd.value, "","SNM") & "' "
	
	Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","����","X")
		
		frm1.txtPlantCd.focus 
		frm1.txtPlantNm.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	frm1.txtPlantNm.value = strDataNm(0)
	

End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery()   

	Dim strVal
			
	DbQuery = False
	    
	If LayerShowHide(1) = False then
		Exit Function 
	End if
	    
	Err.Clear
	
	With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       
		strVal = strVal & "&cboApType=" & Trim(.cboApType.value)
		strVal = strVal & "&txtPlantCd1=" & Trim(.txtPlantCd1.value)
		strVal = strVal & "&txtDocumentDt=" & .txtDocumentDt.text
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows		

		Call RunMyBizASP(MyBizASP, strVal) 

	End With

	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()     
	Dim index
	Dim ii
	'-----------------------
	'Reset variables area
	'-----------------------
   If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd1.focus 
	End If
	Set gActiveElement = document.activeElement
	
End Function

'========================================================
'========================================================

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</script>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</head>

<body tabindex="-1" scroll="no">

<form name="frm1" target="MyBizASP" method="POST">
	<table <%=lr_space_type_00%>>
		<tr>
			<td <%=height_type_00%>></td>
		</tr>
		<tr height="23">
			<td width="100%">
			<table <%=lr_space_type_10%>>
				<tr>
					<td width="10">��</td>
					<td class="CLSMTABP">
					<table id="MyTab" cellspacing="0" cellpadding="0">
						<tr>
							<td background="../../../CShared/image/table/seltab_up_bg.gif">
							<img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" class="CLSMTAB"><font color="#FFFFFF">�������Ͻ�����ȸ</font></td>
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right">
							<img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						</tr>
					</table>
					</td>
					<td width="*">��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr height="*">
			<td width="100%" class="Tab11">
			<table <%=lr_space_type_20%>>
				<tr>
					<td <%=height_type_02%> width="100%"></td>
				</tr>
				<tr>
					<td height="20" width="100%"><fieldset class="CLSFLD">
					<table <%=lr_space_type_40%>>
						<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
								<INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm1" SIZE=25 tag="14X"></TD>							 				
								
								<TD CLASS="TD5" NOWRAP></TD>
							    <TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language="javascript" src="./js/s4513ma1_ko441_fpDateTime1_txtDocumentDt.js"></script>	
									</TD>		
								<TD CLASS="TD5" NOWRAP>�з�</TD>
								<TD CLASS="TD6" NOWRAP>
								<SELECT NAME="cboApType" ALT="�з�" STYLE="Width: 100px;" tag="12"></SELECT>
								</TD>
								
							</TR>
					</table>
					</fieldset> </td>
				</tr>
				<tr>
					<td <%=height_type_03%> width="100%"></td>
				</tr>
				<tr>
					<td width="100%" height="*" valign="TOP">
					<table <%=lr_space_type_60%>>
						<tr>
							<td height="100%" width="100%" colspan="4">
							<script language="javascript" src="./js/s4513ma1_ko441_vspdData.js"></script>
							</td>
						</tr>
					</table>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td <%=height_type_01%>></td>
		</tr>
		<tr>						
	
			<td width="100%" height="<%=BizSize%>">
			<iframe name="MyBizASP" src="../../blank.htm" width="100%" height="<%=BizSize%>" frameborder="0" scrolling="NO" noresize framespacing="0" tabindex="-1">
			</iframe></td>
		</tr>
	</table>
	<textarea class="hidden" name="txtSpread" tag="24"></textarea>
	<p id="divTextArea"></p>
	<input type="HIDDEN" name="txtMode" tag="24" tabindex="-1"><input type="HIDDEN" name="txtMaxRows" tag="24" tabindex="-1">	
	<input type="HIDDEN" name="hdnPlant" tag="24" tabindex="-1"><input type="HIDDEN" name="hdnItem" tag="24" tabindex="-1">
	<input type="HIDDEN" name="hdnItemGroup" tag="24" tabindex="-1"><input type="HIDDEN" name="hdntxtDocumentDt" tag="24" tabindex="-1">
   
	
	
</form>
<div id="MousePT" name="MousePT">
	<iframe name="MouseWindow" frameborder="0" scrolling="NO" noresize framespacing="0" width="220" height="41" src="../../inc/cursor.htm"></iframe>
</div>

</body>

</html>

