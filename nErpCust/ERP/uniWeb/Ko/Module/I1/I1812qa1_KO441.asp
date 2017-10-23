<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : inventory Management
'*  2. Function Name        : 
'*  3. Program ID           : i1812qa1_KO441
'*  4. Program Name         : â�� ������Ȳ 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/07/31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Ho Jun
'* 10. Modifier (Last)      : 
'* 11. Comment
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit														'��: indicates that All variables must be declared in advance

<%
EndDate   = GetSvrDate
%>

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID		= "i1812qb1_KO441.asp"                 '��: �����Ͻ� ���� ASP�� 

Dim C_SlCd			'â��

Dim C_ItemAcctCd	'ǰ������ڵ�
Dim C_ItemAcctNm	'ǰ�������

Dim C_ItemCd		'ǰ��
Dim C_ItemNm		'ǰ���
Dim C_BasicUnit		'����
Dim C_Price			'�ܰ�
Dim C_TransStockQty	'�̿����
Dim C_TransStockAmt	'�̿����ݾ�
Dim C_InProdQty		'�����԰�
Dim C_InProdAmt		'�����԰�ݾ�
Dim C_InPurQty		'�����԰�
Dim C_InPurAmt		'�����԰�ݾ�
Dim C_InExecQty		'�����԰�
Dim C_InExecAmt		'�����԰�ݾ�
Dim C_InStockQty	'����̵��԰�
Dim C_InStockAmt	'����̵��԰�ݾ�
Dim C_InSumQty		'�԰��
Dim C_InSumAmt		'�԰��ݾ�
Dim C_OutProdQty	'�������
Dim C_OutProdAmt	'�������ݾ�
Dim C_OutPurQty		'�Ǹ����
Dim C_OutPurAmt		'�Ǹ����ݾ�
Dim C_OutExecQty	'�������
Dim C_OutExecAmt	'�������ݾ�
Dim C_OutStockQty	'����̵����
Dim C_OutStockAmt	'����̵����ݾ�
Dim C_OutSumQty		'����
Dim C_OutSumAmt		'����(�ݾ�)
Dim C_StockQty		'���
Dim C_StockAmt		'���

'--------------- ������ coding part(��������,End)-------------------------------------------------------------

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
                                         '��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ -----
'--------------- ������ coding part(�������,End)------------------------------------------------------------- 

'==========================================  InitComboBox()  ======================================
'	Name : InitComboBox()
'	Description : Init ComboBox
'==================================================================================================
Sub InitComboBox()

	
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgBlnFlgChgValue = False
	IsOpenPop = False
End Sub                          

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

	Dim StartDate
	StartDate = UNIDateAdd("m", -1,"<%=EndDate%>", parent.gServerDateFormat)
	
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtItemCd.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
  	frm1.txtPlantCd.value = lgPLCd
	End If

	'frm1.txtMovFrDt.Text = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtMovFrDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
	frm1.txtMovToDt.Text = UniConvDateAToB("<%=EndDate%>", parent.gServerDateFormat, parent.gDateFormat) 
 	Call ggoOper.FormatDate(frm1.txtMovFrDt, parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtMovToDt, parent.gDateFormat, 2)

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","QA") %>
End Sub


'--------------------------------------------------------------------------------------------------------- 
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtPlantCd.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "�����ڵ�"		
	arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	

End Function

 '------------------------------------------  OpenItemAcct()  --------------------------------------------------
'	Name : OpenItemAcct()
'	Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "ǰ����� �˾�"											' �˾� ��Ī 
	arrParam(1) = "B_MINOR"																' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemAcct.Value)						' Code Condition
	arrParam(3) = ""																			' Name Cindition
	arrParam(4) = "MAJOR_CD = 'P1001'"										' Where Condition
	arrParam(5) = "ǰ�����"			
	
	arrField(0) = "MINOR_CD"															' Field��(0)
	arrField(1) = "MINOR_NM"															' Field��(1)
	
	arrHeader(0) = "ǰ������ڵ�"															' Header��(0)
	arrHeader(1) = "ǰ�������"																' Header��(1)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemAcct(arrRet)
	End If	
	
End Function

 '------------------------------------------  OpenItemCd()  --------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Cd
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X","X","X")   <% '���������� �ʿ��մϴ� %>
		Exit Function
	End If
	
	'ǰ������ڵ尡 �ִ� �� üũ 
	'If Trim(frm1.txtItemAcct.Value) = "" then
	'	Call DisplayMsgBox("169953","X","X","X")  
	'	Exit Function
	'End If              'update 2002/08/08 lsw

	'�����ڵ� �� ǰ������ڵ� üũ �Լ� ȣ�� 
	'If Plant_Or_ItemAcct_Check = False Then 
	'	Exit Function
	'End If

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "ǰ�� �˾�"											' �˾� ��Ī 
	arrParam(1) = "B_ITEM_BY_PLANT P,B_ITEM I"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemCd.Value)				        		' Code Condition
	arrParam(3) = ""																			' Name Cindition
	'arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value,"","S") & " AND P.ITEM_ACCT=" & FilterVar(frm1.txtItemAcct.Value,"","S") ' Where Condition
	arrParam(4) = "P.ITEM_CD=I.ITEM_CD AND P.PLANT_CD =" & FilterVar(frm1.txtPlantCd.Value,"","S") ' Where Condition
	arrParam(5) = "ǰ��"			
	
	arrField(0) = "I.ITEM_CD"															' Field��(0)
	arrField(1) = "I.ITEM_NM"															' Field��(1)
	
	arrHeader(0) = "ǰ���ڵ�"															' Header��(0)
	arrHeader(1) = "ǰ���"																' Header��(1)
	
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
'	lgBlnFlgChgValue	  	 = True	
End Function

 '------------------------------------------  SetItemAcct()  --------------------------------------------------
'	Name : SetItemAcct()
'	Description : ItemAcct Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAcct(byval arrRet)
	frm1.txtItemAcct.Value	    =arrRet(0)
	frm1.txtItemAcctNm.Value	=arrRet(1)
	frm1.txtItemAcct.focus
	Set gActiveElement = document.activeElement
End Function

 '------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : ItemAcct Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value	=arrRet(0)
	frm1.txtItemNm.Value	=arrRet(1)
	frm1.txtItemAcct.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenSL()  --------------------------------------------------
'	Name : OpenSL()
'	Description : SL Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	'-----------------------
	'Check Plant CODE	
	'-----------------------
	'If	CommonQueryRs("	PLANT_NM "," B_PLANT ",	" PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
	'	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
		
	'	Call DisplayMsgBox("125000","X","X","X")
	'	frm1.txtPlantNm.Value = ""
	'	frm1.txtPlantCd.focus
	'	Exit function
	'End	If
	'lgF0 = Split(lgF0, Chr(11))
	'frm1.txtPlantNm.Value = lgF0(0)

	IsOpenPop = True

	arrParam(0) = "â���˾�"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	if frm1.txtPlantCd.value <> "" then
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value,"","S")	
	else
	arrParam(4) = ""
	end if
	arrParam(5) = "â��"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"
	
	arrHeader(0) = "â��"		
	arrHeader(1) = "â���"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
End Function


'------------------------------------------  SetSL()  --------------------------------------------------
'	Name : SetSL()
'	Description : SL Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSL(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
'	lgBlnFlgChgValue	  = True
End Function

'========================================================================================
' Function Name : Plant_Or_ItemAcct_Check
' Function Desc : 
'========================================================================================
Function Plant_Or_ItemAcct_Check()
	'-----------------------
	'Check Plant CODE		'�����ڵ尡 �ִ� �� üũ 
	'-----------------------
    'If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
	'	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	'	
	'	Call DisplayMsgBox("125000","X","X","X")
	'	frm1.txtPlantNm.Value = ""
	'	frm1.txtPlantCd.focus
	'	Plant_Or_ItemAcct_Check = False
	'	Exit function
	'	
    'End If
	'lgF0 = Split(lgF0, Chr(11))
	'frm1.txtPlantNm.Value = lgF0(0)

	'-----------------------
	'Check ItemAcct CODE	''ǰ������ڵ尡 �ִ� �� üũ 
	'-----------------------
    If 	CommonQueryRs(" MINOR_NM "," B_MINOR ", " MAJOR_CD = 'P1001' AND MINOR_CD= " & FilterVar(frm1.txtItemAcct.Value,"","S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		
		Call DisplayMsgBox("169952","X","X","X")
		frm1.txtItemAcctNm.Value = ""
		frm1.txtItemAcct.focus
		Plant_Or_ItemAcct_Check = False
		Exit function
    End If
    lgF0 = Split(lgF0, Chr(11))
	frm1.txtItemAcctNm.Value = lgF0(0)

	'-----------------------
	'Check Item CODE		'ǰ��� ��ȸ 
	'-----------------------
    If frm1.txtItemCd.Value <> "" Then
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value,"","S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm.Value = lgF0(0)
		Else
			frm1.txtItemNm.Value = ""
		End If
	End If

    Plant_Or_ItemAcct_Check = True
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()

	on error resume next
	
	Dim Ret
	Call InitSpreadPosVariables()
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030804", , Parent.gAllowDragDropSpread

	With frm1.vspdData
	
		.ReDraw = false
		
		.MaxCols = C_StockAmt + 1
		.ColHeaderRows = 2    '����� 2�ٷ�
		.Col = .MaxCols
		.MaxRows = 0
		
 		Call GetSpreadColumnPos("A")

'2008-05-22 1:06���� :: hanc :: 12->15 &  	ggAmtOfMoneyNo->ggQtyNo 	
 		ggoSpread.SSSetEdit	  C_SlCd,			"â���ڵ�",		10
 		ggoSpread.SSSetEdit	  C_ItemAcctCd,		"ǰ�����",		10
 		ggoSpread.SSSetEdit	  C_ItemAcctNm,		"ǰ�������",	10
		ggoSpread.SSSetEdit	  C_ItemCd,			"ǰ���ڵ�",		15
		ggoSpread.SSSetEdit   C_ItemNm,			"ǰ���",		20
		ggoSpread.SSSetEdit   C_BasicUnit,		"����",			8
		ggoSpread.SSSetFloat  C_Price,			"���شܰ�",		10, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_TransStockQty,	"�̿����",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_TransStockAmt,	"�̿����ݾ�",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InProdQty,		"�����԰����",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InProdAmt,		"�����԰�ݾ�",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InPurQty,		"�����԰����",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InPurAmt,		"�����԰�ݾ�",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InExecQty,		"�����԰�(�԰�)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InExecAmt,		"�����԰�ݾ�(�԰�)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InStockQty,		"����̵�(�԰�)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InStockAmt,		"����̵��ݾ�(�԰�)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InSumQty,		"��(�԰�)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_InSumAmt,		"��ݾ�(�԰�)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutProdQty,		"�������(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutProdAmt,		"�������ݾ�(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutPurQty,		"�Ǹ����(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutPurAmt,		"�Ǹ����ݾ�(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutExecQty,		"�������(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutExecAmt,		"�������ݾ�(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutStockQty,	"����̵�(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutStockAmt,	"����̵��ݾ�(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutSumQty,		"��(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_OutSumAmt,		"��ݾ�(���)",	15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_StockQty,		"���",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		ggoSpread.SSSetFloat  C_StockAmt,		"���ݾ�",		15, Parent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
'		ggoSpread.SSSetFloat  C_Price,			"���شܰ�",		10, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_TransStockQty,	"�̿����",		12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_TransStockAmt,	"�̿����ݾ�",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InProdQty,		"�����԰����",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InProdAmt,		"�����԰�ݾ�",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InPurQty,		"�����԰����",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InPurAmt,		"�����԰�ݾ�",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InExecQty,		"�����԰�(�԰�)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InExecAmt,		"�����԰�ݾ�(�԰�)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InStockQty,		"����̵�(�԰�)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InStockAmt,		"����̵��ݾ�(�԰�)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_InSumQty,		"��(�԰�)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_InSumAmt,		"��ݾ�(�԰�)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutProdQty,		"�������(���)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutProdAmt,		"�������ݾ�(���)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutPurQty,		"�Ǹ����(���)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutPurAmt,		"�Ǹ����ݾ�(���)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutExecQty,		"�������(���)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutExecAmt,		"�������ݾ�(���)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutStockQty,	"����̵�(���)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutStockAmt,	"����̵��ݾ�(���)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_OutSumQty,		"��(���)",	12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
'		ggoSpread.SSSetFloat  C_OutSumAmt,		"��ݾ�(���)",	12, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
'		ggoSpread.SSSetFloat  C_StockQty,		"���",		12, Parent.ggQty, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,Z
		
		
		' �׸��� ��� ��ħ 
		ret = .AddCellSpan(C_SlCd			, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemAcctCd		, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemAcctNm		, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemCd			, -1000, 1, 2)
		ret = .AddCellSpan(C_ItemNm			, -1000, 1, 2)
		ret = .AddCellSpan(C_BasicUnit		, -1000, 1, 2)
		ret = .AddCellSpan(C_Price			, -1000, 1, 2)
		ret = .AddCellSpan(C_TransStockQty	, -1000, 1, 2)
		ret = .AddCellSpan(C_TransStockAmt	, -1000, 1, 2)
		
		ret = .AddCellSpan(C_InProdQty		, -1000, 10, 1)
		ret = .AddCellSpan(C_OutProdQty		, -1000, 10, 1)
		ret = .AddCellSpan(C_StockQty		, -1000, 2, 1)
			
		' ù��° ��� ��� ���� 
		.Row = -1000
		.Col = C_SlCd			: .Text = "â���ڵ�"
		.Col = C_ItemAcctCd		: .Text = "ǰ�����"
		.Col = C_ItemAcctNm		: .Text = "ǰ�������"
		.Col = C_ItemCd			: .Text = "ǰ���ڵ�"
		.Col = C_ItemNm			: .Text = "ǰ���"
		.Col = C_BasicUnit		: .Text = "����"
		.Col = C_Price			: .Text = "���شܰ�"
		.Col = C_TransStockQty	: .Text = "�̿����"
		.Col = C_TransStockAmt	: .Text = "�̿����ݾ�"
		.Col = C_InProdQty		: .Text = "�԰�"
		.Col = C_OutProdQty		: .Text = "���"
		.Col = C_StockQty		: .Text = "���"
			
		.Row = -999
		
		.Col = C_SlCd			: .Text = "â���ڵ�"
		.Col = C_ItemAcctCd		: .Text = "ǰ�����"
		.Col = C_ItemAcctNm		: .Text = "ǰ�������"
		.Col = C_ItemCd			: .Text = "ǰ���ڵ�"
		.Col = C_ItemNm			: .Text = "ǰ���"
		.Col = C_BasicUnit		: .Text = "����"
		.Col = C_Price			: .Text = "���شܰ�"	
		.Col = C_TransStockQty	: .Text = "�̿����"
		.Col = C_TransStockAmt	: .Text = "�̿����ݾ�"
		.Col = C_InProdQty		: .Text = "�����԰����"
		.Col = C_InProdAmt		: .Text = "�����԰�ݾ�"
		.Col = C_InPurQty		: .Text = "�����԰����"
		.Col = C_InPurAmt		: .Text = "�����԰�ݾ�"
		.Col = C_InExecQty		: .Text = "�����԰�"
		.Col = C_InExecAmt		: .Text = "�����԰�ݾ�"
		.Col = C_InStockQty		: .Text = "����̵�"
		.Col = C_InStockAmt		: .Text = "����̵��ݾ�"
		.Col = C_InSumQty		: .Text = "��"
		.Col = C_InSumAmt		: .Text = "��ݾ�"
		.Col = C_OutProdQty		: .Text = "�������"
		.Col = C_OutProdAmt		: .Text = "�������ݾ�"
		.Col = C_OutPurQty		: .Text = "�Ǹ����" 				
		.Col = C_OutPurAmt		: .Text = "�Ǹ����ݾ�" 					
		.Col = C_OutExecQty		: .Text = "�������" 					
		.Col = C_OutExecAmt		: .Text = "�������ݾ�" 					
		.Col = C_OutStockQty	: .Text = "����̵�" 			
		.Col = C_OutStockAmt	: .Text = "����̵��ݾ�"
		.Col = C_OutSumQty		: .Text = "��"
		.Col = C_OutSumAmt		: .Text = "��ݾ�"
		.Col = C_StockQty		: .Text = "���"
		.Col = C_StockAmt		: .Text = "���ݾ�"
		
		.RowHeight(-999) = 15	' ���� ������	(2���� ���, 1���� 15)	
 	
 		Call ggoSpread.SSSetColHidden(C_Price, C_Price, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    'ggoSpread.SSSetSplit2(2)  
	    'Call SetSpreadLock()
		
		.ReDraw = true
		
    End With
End Sub

Sub SetSpreadLock()
   
End Sub

'==========================================  2.6.1 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()

	C_SlCd			= 1			'â��
	C_ItemAcctCd	= 2			'ǰ�����
	C_ItemAcctNm	= 3			'ǰ�������
	C_ItemCd		= 4			'ǰ��
	C_ItemNm		= 5			'ǰ���
	C_BasicUnit		= 6			'����
	C_Price			= 7			'�ܰ�
	C_TransStockQty	= 8			'�̿����
	C_TransStockAmt	= 9			'�̿����ݾ�
	C_InProdQty		= 10		'�����԰�
	C_InProdAmt		= 11		'�����԰�ݾ�
	C_InPurQty		= 12		'�����԰�
	C_InPurAmt		= 13		'�����԰�ݾ�
	C_InExecQty		= 14		'�����԰�
	C_InExecAmt		= 15		'�����԰�ݾ�
	C_InStockQty	= 16		'����̵��԰�
	C_InStockAmt	= 17		'����̵��԰�ݾ�
	C_InSumQty		= 18		'���԰�
	C_InSumAmt		= 19		'���԰�ݾ�
	C_OutProdQty	= 20		'�������
	C_OutProdAmt	= 21		'�������ݾ�
	C_OutPurQty		= 22		'�Ǹ����
	C_OutPurAmt		= 23		'�Ǹ����ݾ�
	C_OutExecQty	= 24		'�������
	C_OutExecAmt	= 25		'�������ݾ�
	C_OutStockQty	= 26		'����̵����
	C_OutStockAmt	= 27		'����̵����ݾ�
	C_OutSumQty		= 28		'�����
	C_OutSumAmt		= 29		'�����ݾ�
	C_StockQty		= 30		'���
	C_StockAmt		= 31		'���ݾ�

End Sub

'==========================================  2.6.2 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_SlCd			= iCurColumnPos(1)		'â���ڵ�
		C_ItemAcctCd	= iCurColumnPos(2)		'ǰ�����
		C_ItemAcctNm	= iCurColumnPos(3)		'ǰ�������
		C_ItemCd		= iCurColumnPos(4)		'ǰ��
		C_ItemNm		= iCurColumnPos(5)		'ǰ���
		C_BasicUnit		= iCurColumnPos(6)		'����
		C_Price			= iCurColumnPos(7)		'�ܰ�
		C_TransStockQty	= iCurColumnPos(8)		'�̿����
		C_TransStockAmt	= iCurColumnPos(9)		'�̿����ݾ�
		C_InProdQty		= iCurColumnPos(10)		'�����԰�
		C_InProdAmt		= iCurColumnPos(11)		'�����԰�ݾ�
		C_InPurQty		= iCurColumnPos(12)		'�����԰�
		C_InPurAmt		= iCurColumnPos(13)		'�����԰�ݾ�
		C_InExecQty		= iCurColumnPos(14)		'�����԰�
		C_InExecAmt		= iCurColumnPos(15)		'�����԰�ݾ�
		C_InStockQty	= iCurColumnPos(16)		'����̵��԰�
		C_InStockAmt	= iCurColumnPos(17)		'����̵��԰�ݾ�
		C_InSumQty		= iCurColumnPos(18)		'���԰�
		C_InSumAmt		= iCurColumnPos(19)		'���԰�ݾ�
		C_OutProdQty	= iCurColumnPos(20)		'�������
		C_OutProdAmt	= iCurColumnPos(21)		'�������ݾ�
		C_OutPurQty		= iCurColumnPos(22)		'�Ǹ����
		C_OutPurAmt		= iCurColumnPos(23)		'�Ǹ����ݾ�
		C_OutExecQty	= iCurColumnPos(24)		'�������
		C_OutExecAmt	= iCurColumnPos(25)		'�������ݾ�
		C_OutStockQty	= iCurColumnPos(26)		'����̵����
		C_OutStockAmt	= iCurColumnPos(27)		'����̵����ݾ�
		C_OutSumQty		= iCurColumnPos(28)		'�����
		C_OutSumAmt		= iCurColumnPos(29)		'�����ݾ�
		C_StockQty		= iCurColumnPos(30)		'���
		C_StockAmt		= iCurColumnPos(31)		'���ݾ�
				
 	End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

	Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	
	Call InitVariables														'��: Initializes local global variables
  Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitComboBox()
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'��: ��ư ���� ����	
'--------------- ������ coding part(�������,Start)----------------------------------------------------
   	
	frm1.txtIssueReqNo.focus
'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col				
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		
 			lgSortKey = 1
 		End If
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

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
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
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then Exit Sub
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If    
End Sub



'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then Exit Function								'��: This function check indispensable field
    
   	
    '-----------------------
    'Erase contents area
    '-----------------------
    ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 

	'If Name_check("A") = False Then
	'	Set gActiveElement = document.activeElement
	'	Exit Function
	'End If
	
    Call InitVariables 	
    
    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then Exit Function

    FncQuery = True															'��: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

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
	Dim strVal
	Dim rdoIssue
	Dim rdoConfirm
	Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
	
	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtPlantCd="		& Trim(.txtPlantCd.value) & _							
							  "&txtSlCd=" & Trim(.txtSLCd.value) & _
							  "&txtFromYY=" & Trim(strYear) & _
							  "&txtFromMm=" & Trim(strMonth) & _
							  "&txtToYY=" & Trim(strYear1) & _
							  "&txtToMm=" & Trim(strMonth1) & _		
							  "&txtItemAcct=" & Trim(.txtItemAcct.value) & _	
							  "&txtItemCd=" & Trim(.txtItemCd.value) & _		
							  "&txtMaxRows="	& .vspdData.MaxRows & _
							  "&lgStrPrevKey="	& lgStrPrevKey                      '��: Next key tag
					  
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
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
    Call SetToolbar("11000000000111")
	lgBlnFlgChgValue = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview()
    Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
	Dim var1, var2, var3, var4, var5, var6, var7, var8
    Dim condvar 
 	Dim ObjName
   
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then                             '��: Check contents area
       Exit Function
    End If
    
    '�����ڵ� �� ǰ������ڵ� üũ �Լ� ȣ�� 
    'If Plant_Or_ItemAcct_Check = False Then 
	'	Exit Function
	'End If
	
    If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
	
 	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    
    var1 = UCASE(Trim(frm1.txtPlantCd.value))
    If var1 = "" Then
       var1 = "%"
    End If
    var2 = Trim(frm1.txtItemCd.Value)
    var3 = frm1.txtItemAcct.Value
    var4 = strYear
    var5 = strMonth
    var6 = strYear1
    var7 = strMonth1
    var8 = Trim(frm1.txtSLCd.value)
    
	If var2 = "" Then
       var2 = "%"
    End If
    
	If var3 = "" Then
       var3 = "%"
    End If    
    
	If var8 = "" Then
       var8 = "%"
    End If

	ObjName = AskEBDocumentName("I1812OA2_KO441", "ebr")				'â��
		
	condvar = condvar & "PLANTCD|"    & var1
	condvar = condvar & "|SLCD|"      & var8
    condvar = condvar & "|ITEMCD|"    & var2
    condvar = condvar & "|ITEMACCT|"  & var3
    condvar = condvar & "|FromYy|"    & var4
    condvar = condvar & "|FromMm|"    & var5
    condvar = condvar & "|ToYy|"      & var6
    condvar = condvar & "|ToMm|"      & var7

	Call FncEBRPreview(ObjName, condvar)    
	
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 
    Dim strYear, strMonth, strDay, strYear1, strMonth1, strDay1
	Dim var1, var2, var3, var4, var5, var6, var7, var8
    Dim condvar 
 	Dim ObjName
   
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then                             '��: Check contents area
       Exit Function
    End If
    
    '�����ڵ� �� ǰ������ڵ� üũ �Լ� ȣ�� 
    'If Plant_Or_ItemAcct_Check = False Then 
	'	Exit Function
	'End If
	
    If ValidDateCheck(frm1.txtMovFrDt, frm1.txtMovToDt) = False Then
		Exit Function
    End If
	
 	Call ExtractDateFrom(frm1.txtMovFrDt.Text,frm1.txtMovFrDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
  	Call ExtractDateFrom(frm1.txtMovToDt.Text,frm1.txtMovToDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)
    
    var1 = UCASE(Trim(frm1.txtPlantCd.value))
    If var1 = "" Then
       var1 = "%"
    End If
    var2 = Trim(frm1.txtItemCd.Value)
    var3 = frm1.txtItemAcct.Value
    var4 = strYear
    var5 = strMonth
    var6 = strYear1
    var7 = strMonth1
    var8 = Trim(frm1.txtSLCd.value)
    
	If var2 = "" Then
       var2 = "%"
    End If
    
    If var3 = "" Then
       var3 = "%"
    End If  
    
	If var8 = "" Then
       var8 = "%"
    End If
     
    ObjName = AskEBDocumentName("I1812OA1_KO441", "ebr")				'â��
    
	condvar = condvar & "PLANTCD|"    & var1
	condvar = condvar & "|SLCD|"      & var8
	condvar = condvar & "|PLANTCD|"   & var1
    condvar = condvar & "|ITEMCD|"    & var2
    condvar = condvar & "|ITEMACCT|"  & var3
    condvar = condvar & "|FromYy|"    & var4
    condvar = condvar & "|FromMm|"    & var5
    condvar = condvar & "|ToYy|"      & var6
    condvar = condvar & "|ToMm|"      & var7
    
	Call FncEBRprint(EBAction, ObjName, condvar) 	

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>â��������Ȳ��ȸ(S)</font></td>
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
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>									
        							<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6">
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=20 tag="14">
									</TD>
									<TD CLASS="TD5" NOWRAP>���ұⰣ</TD>
									<TD CLASS="TD6">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtMovFrDt classid=<%=gCLSIDFPDT%> tag="12x1" ALT="�˻����۳�¥" VIEWASTEXT><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</Script>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtMovToDt classid=<%=gCLSIDFPDT%> ALT="�˻�����¥" tag="12x1" VIEWASTEXT><PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""></OBJECT>');</Script>
									</TD>
								</TR>
								<TR>									
        							<TD CLASS="TD5">ǰ�����</TD>
									<TD CLASS="TD6">
									<input TYPE=TEXT NAME="txtItemAcct" SIZE="8" MAXLENGTH="2" tag="11XXXU" ALT="ǰ�����"  ><IMG align=top height=20 name="btnItemAcct" onclick="vbscript:OpenItemAcct()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 MAXLENGTH=40 tag="14">
									</TD>
									<TD CLASS="TD5">ǰ��</TD>
									<TD CLASS="TD6">
									<input TYPE=TEXT NAME="txtItemCd" SIZE="18" MAXLENGTH="18" ALT="ǰ��" tag="11XXXU" ><IMG align=top height=20 name="btnItemNm" onclick="vbscript:OpenItemCd()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=40 tag="14">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>â��</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE="TEXT" NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="11XXXU" ALT="â��" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE="TEXT" NAME="txtSLNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="â���">
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><TD>
									</TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TR>
							<TD HEIGHT=100% WIDTH=100% Colspan=2>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>	
						</TR>	
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreView()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>
	<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname">
	<input type="hidden" name="dbname">
	<input type="hidden" name="filename">
	<input type="hidden" name="condvar">
	<input type="hidden" name="date">                 
	</FORM>
</BODY>
</HTML>
