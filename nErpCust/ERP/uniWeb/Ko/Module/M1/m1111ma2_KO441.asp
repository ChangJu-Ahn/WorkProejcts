<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M1111MA1
'*  4. Program Name         : ǰ�񺰴ܰ���� 
'*  5. Program Desc         : ǰ�񺰴ܰ���� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/08
'*  8. Modified date(Last)  : 2005/06/30
'*  9. Modifier (First)     : Sin jin hyun
'* 10. Modifier (Last)      : Yim Young Ju
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
<!-- '******************************************  1.1 Inc ����   **********************************************
' ���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
<!-- #Include file="../../inc/lgvariables.inc" -->

<!--'******************************************  1.2 Global ����/��� ����  ***********************************
' 1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************!-->
Const BIZ_PGM_ID = "m1111mb2_ko441.asp"

<!--'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================!-->
Dim lgIsOpenPop
Dim C_PlantCd
Dim C_PlantPopup
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemPopup
Dim C_ItemNm
DIM C_SPEC

Dim C_OrderUnit
Dim C_OrderUnitPopup
Dim C_Curr
Dim C_CurrPopUP
Dim C_AppDt
Dim C_Cost
'�̼��� �߰� 
Dim C_PrcFlg
Dim C_PrcFlgNm
'2005.05.03 �߰� 
Dim C_Remark

<!-- '==========================================  1.2.2 Global ���� ����  =====================================
' 1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
' 2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= !-->
Dim iDBSYSDate
Dim EndDate, StartDate

	'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
	EndDate = "<%=GetSvrDate%>"
	'------ ��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ ------
	StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
	EndDate = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
	StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

<!-- '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= !-->
Dim IsOpenPop          

<!-- '==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= !-->
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE  
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0          
    
    lgStrPrevKey = ""          
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
End Sub
 
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()

	C_PlantCd			=	1
	C_PlantPopup		=	2
	C_PlantNm			=	3
	C_ItemCd			=	4
	C_ItemPopup			=	5
	C_ItemNm			=	6
	C_SPEC				 =  7
	  
	C_OrderUnit		 =	8
	C_OrderUnitPopup = 9
	C_Curr			 = 10
	C_CurrPopUP		 = 11
	C_AppDt			 = 12
	C_Cost			 = 13
	C_PrcFlg		 = 14
	C_PrcFlgNm		 = 15
	C_Remark		= 16
End Sub

<!-- '==========================================  2.2.1 SetDefaultVal()  ========================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= !-->
 Sub SetDefaultVal()
	Call SetToolBar("1110110100101111")
	frm1.txtPlantCd1.value=parent.gPlant
	frm1.txtPlantNm1.value=parent.gPlantNm 
	frm1.txtPlantCd1.focus 

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd1, "Q") 
  	frm1.txtPlantCd1.value = lgPLCd
	End If


	Set gActiveElement = document.activeElement
End Sub

<!--'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== !-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

<!--
'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
-->
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
	 
	With frm1.vspdData
		.ReDraw = False
		.MaxCols = C_Remark+1
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit   C_PlantCd, "����",  10,,,,2
		ggoSpread.SSSetButton C_PlantPopup
		ggoSpread.SSSetEdit   C_PlantNm, "�����", 20		
		ggoSpread.SSSetEdit   C_ItemCd, "ǰ��", 15,,,,2
		ggoSpread.SSSetButton C_ItemPopup
		ggoSpread.SSSetEdit   C_ItemNm, "ǰ���", 20
		ggoSpread.SSSetEdit   C_SPEC, "�԰�", 20				
		 
		ggoSpread.SSSetEdit   C_OrderUnit, "���ִ���", 8,2,,3,2
		ggoSpread.SSSetButton C_OrderUnitPopup
		ggoSpread.SSSetEdit   C_Curr, "ȭ��", 8,2,,3,2
		ggoSpread.SSSetButton C_CurrPopUP
		ggoSpread.SSSetDate   C_AppDt, "�ܰ�������", 10,2,gDateFormat
        ggoSpread.SSSetFloat  C_Cost , "�ܰ�"       ,15    ,"C"   ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, 1,,"Z"
        ggoSpread.SSSetCombo  C_PrcFlg , "�ܰ�����" , 10
        ggoSpread.SSSetCombo  C_PrcFlgNm , "�ܰ�����" , 10  
		ggoSpread.SSSetEdit   C_Remark, "���", 50,0,,240,0
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.MakePairsColumn(C_OrderUnit,C_OrderUnitPopup)
		Call ggoSpread.MakePairsColumn(C_Curr,C_CurrPopup)
		
		.ReDraw = true
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
            
            C_PlantCd			=	iCurColumnPos(1)
			C_PlantPopup		=	iCurColumnPos(2)
			C_PlantNm			=	iCurColumnPos(3)
			C_ItemCd			=	iCurColumnPos(4)	
			C_ItemPopup			=	iCurColumnPos(5)
			C_ItemNm			=	iCurColumnPos(6)
			C_SPEC				=	iCurColumnPos(7)
			            
			C_OrderUnit		 = iCurColumnPos(8)   
			C_OrderUnitPopup = iCurColumnPos(9)
			C_Curr			 = iCurColumnPos(10)
			C_CurrPopUP		 = iCurColumnPos(11)   
			C_AppDt			 = iCurColumnPos(12)
			C_Cost			 = iCurColumnPos(13)
			'20050504 �߰� 
			C_PrcFlg		= iCurColumnPos(14)
			C_PrcFlgNm		= iCurColumnPos(15)
			C_Remark		= iCurColumnPos(16)
			
    End Select    
End Sub

<!--
'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
-->
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_Orderunit, -1
    ggoSpread.spreadlock C_Orderunitpopup, -1
    ggoSpread.spreadlock C_Curr, -1
    ggoSpread.spreadlock C_CurrPopup, -1
    ggoSpread.spreadlock C_AppDt, -1
    ggoSpread.spreadUnlock C_Cost, -1
    ggoSpread.sssetrequired C_Cost, -1
    
    ggoSpread.spreadlock C_PrcFlgNm, -1    
    
    ggoSpread.SSSetProtected frm1.vspdData.MaxCols, -1
    .vspdData.ReDraw = True
    End With
End Sub

<!--
'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
-->

Sub SetSpreadColor1(ByVal pvStarRow, Byval pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With frm1
    	.vspdData.ReDraw = False
    	
    	ggoSpread.SSSetRequired  C_PlantCd,			pvStarRow,	pvEndRow
		ggoSpread.SSSetProtected C_PlantNm,			pvStarRow,	pvEndRow
		ggoSpread.SSSetRequired  C_ItemCd,			pvStarRow,	pvEndRow
		ggoSpread.SSSetProtected C_ItemNm,			pvStarRow,	pvEndRow
		ggoSpread.SSSetProtected C_SPEC,			pvStarRow,	pvEndRow
		    	
		ggoSpread.SSSetRequired  C_OrderUnit,		pvStarRow, pvEndRow
		ggoSpread.SSSetRequired  C_Curr,			pvStarRow, pvEndRow
		ggoSpread.SSSetRequired  C_AppDt,			pvStarRow, pvEndRow
		ggoSpread.SSSetRequired  C_Cost,			pvStarRow, pvEndRow
		'�̼��� �߰� 
		ggoSpread.SSSetRequired  C_PrcFlg,			pvStarRow, pvEndRow
		ggoSpread.SSSetProtected  C_PrcFlgNm,		pvStarRow, pvEndRow
		
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, pvStarRow, pvEndRow

		if lgPLCd <> "" then 
			ggoSpread.SSSetProtected C_PlantCd,			pvStarRow,	pvEndRow
			ggoSpread.SSSetProtected C_PlantPopup,			pvStarRow,	pvEndRow
		End If
		
		
		.vspdData.ReDraw = True
    End With
End Sub


'�̼��� �߰� 
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim strCboCd
	Dim strCboNm
	
	strCboCd = "" & "T" & vbTab & "F"
	ggoSpread.SetCombo strCboCd, C_PrcFlg  	
	
	strCboNm = "" & "���ܰ�" & vbTab & "���ܰ�"
	ggoSpread.SetCombo strCboNm, C_PrcFlgNm  	  		
End Sub

'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	For intRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = intRow

		frm1.vspdData.Col = C_PrcFlg
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_PrcFlgNm
		frm1.vspdData.value = intindex
	Next
End Sub
	
<!-- '------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenPlant(byval strComp)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	'If IsOpenPop = True Or UCase(frm1.txtPlantCd2.className) = UCase(parent.UCN_PROTECTED) And strComp <> "Plant1"  Then Exit Function
	If frm1.txtPlantCd1.className = "protected" Then Exit Function
	 
	IsOpenPop = True

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
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtPlantCd1.focus
		Exit Function
	Else
		frm1.txtPlantCd1.Value  = arrRet(0)  
		frm1.txtPlantNm1.Value  = arrRet(1)  
		frm1.txtPlantCd1.focus
	End If 
End Function

'------------------------------------------  OpenSPlant()  -------------------------------------------------
' Name : OpenSPlant()
' Description : SpreadPlant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSPlant(byval strCon)
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "�����˾�" 
	arrParam(1) = "B_Plant"    
	arrParam(2) = Trim(strCon)
	arrParam(4) = ""   
	arrParam(5) = "����"   
	 
	arrField(0) = "Plant_Cd" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "����"  
	arrHeader(1) = "�����"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False

	If arrRet(0) = "" Then 
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_PlantCd
			.text = arrRet(0) 
			.Col = C_PlantNm
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
			.focus
		End With 
	End If 
End Function

'------------------------------------------  OpenSItem()  -------------------------------------------------
' Name : OpenSItem()
' Description : SpreadItem PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSItem(byval strCon)
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	IsOpenPop = True
	
	With frm1.vspdData
	
		.Row = .ActiveRow 
	
		arrParam(0) = "ǰ���˾�"      
		arrParam(1) = "B_Item_By_Plant,	B_Item"
		arrParam(2) = Trim(strCon)
		arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
		 
		.Col = C_PlantCd
		If Trim(.text) <> "" then
			arrParam(4) = arrParam(4) & "And B_Item_By_Plant.Plant_Cd= " & FilterVar(.text, "''", "S")
		End If 

		arrParam(5) = "ǰ��"       

		arrField(0) = "B_Item_By_Plant.Item_Cd"     
		arrField(1) = "B_Item.Item_NM" 
	
	End With
	
	iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_ItemNm
			.text = arrRet(1)
			.Col = C_ItemCd
			.text = arrRet(0)
			.Col = C_SPEC
			.text = arrRet(2)  
			Call SetFocusToDocument("M") 
			.focus
		End With 
	End If 
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Plant PopUp ǰ�� 
'===================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd1.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd1.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd1.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd1.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd1.focus
		Exit Function
	Else
		frm1.txtItemCd1.Value	= arrRet(0)
		frm1.txtItemNm1.Value	= arrRet(1)
		frm1.txtItemCd1.focus
	End If
End Function


Function OpenUnit(byval strCon)  
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ִ���"     
	arrParam(1) = "B_Unit_OF_MEASURE"   
	 
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_OrderUnit
	arrParam(2) = Trim(frm1.vspdData.text) 
	 
	arrParam(4) = ""      
	arrParam(5) = "���ִ���"    
	 
	arrField(0) = "Unit"     
	arrField(1) = "Unit_Nm"     
	    
	arrHeader(0) = "���ִ���"   
	arrHeader(1) = "���ִ�����"   
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Row = frm1.vspdData.ActiveRow 
		frm1.vspdData.Col = C_OrderUnit
		frm1.vspdData.text = arrRet(0) 
	End If 
End Function


Function OpenCurr(byval strCon) 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ȭ��"      
	arrParam(1) = "B_Currency"     
	 
	frm1.vspdData.Row = frm1.vspdData.ActiveRow 
	frm1.vspdData.Col = C_Curr
	arrParam(2) = Trim(frm1.vspdData.text) 
	arrParam(3) = ""
	  
	arrParam(4) = ""      
	arrParam(5) = "ȭ��"     
	 
	arrField(0) = "Currency"    
	arrField(1) = "Currency_Desc"   
	    
	arrHeader(0) = "ȭ��"    
	arrHeader(1) = "ȭ���"    
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Row    = frm1.vspdData.ActiveRow 
		frm1.vspdData.Col    = C_Curr
		frm1.vspdData.text    = arrRet(0) 
		Call vspdData_Change(C_Curr,frm1.vspdData.ActiveRow)
	End If 
End Function

<!--
'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : ���Ÿ� ���� �׸����� ���� �κ��� ����ȸ� �� �Լ��� ���� �ؾ���.
'==========================================================================================
-->
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
    End Select         
End Sub

<!--
'++++++++++++++++++++++++++++++++++++++++++++  SetSpreadLockAfterQuery()  +++++++++++++++++++++++++++++++++++++++++
'+ Name : SetSpreadLockAfterQuery()                    +
'+ Description : Set Return array from bank_acct PopUp Window           +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
-->
Sub SetSpreadLockAfterQuery()
	Dim index 

	With frm1
		.vspdData.ReDraw = False
		For index = Cint(.hdnmaxrow.value) + 1 to .vspdData.MaxRows 
			.vspdData.Row = index
			
			ggoSpread.SpreadLock C_PlantCd, index, C_PlantCd, index
			ggoSpread.SpreadLock C_PlantPopup, index, C_PlantPopup, index
			ggoSpread.SpreadLock C_PlantNm, index, C_PlantNm, index 
			ggoSpread.SpreadLock C_ItemCd, index, C_ItemCd, index 
			ggoSpread.SpreadLock C_ItemPopup, index, C_ItemPopup, index 
			ggoSpread.SpreadLock C_ItemNm, index, C_ItemNm, index
			ggoSpread.SpreadLock C_SPEC, index, C_ItemNm, index
						
			ggoSpread.SpreadLock C_Orderunit, index, C_Orderunit, index
			ggoSpread.SpreadLock C_Orderunitpopup, index, C_Orderunitpopup, index
			ggoSpread.SpreadLock C_Curr, index, C_Curr, index
			ggoSpread.SpreadLock C_CurrPopup, index, C_CurrPopup, index
			ggoSpread.SpreadLock C_AppDt, index, C_AppDt, index
			ggoSpread.spreadUnlock C_Cost, index, C_Cost, index
			ggoSpread.SSSetRequired C_Cost, index, index
			'�̼��� �߰� 
			ggoSpread.spreadUnlock C_PrcFlg, index, C_PrcFlg, index
			ggoSpread.SSSetRequired C_PrcFlg, index, index			
			
			ggoSpread.SpreadLock C_PrcFlgNm, index, C_PrcFlgNm, index
		Next
		.vspdData.ReDraw = True
	End With    
End Sub

'========================================================================================
' Function Name : SetRow
' Function Desc : Dev. Define(K.J.H)
'========================================================================================
Function SetRow(iRow)
	Call SetActiveCell(frm1.vspdData,1,iRow,"M","X","X")
End Function

<!-- '==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= !-->
 Sub Form_Load()
    Call LoadInfTB19029                     
    Call ggoOper.LockField(Document, "N")   
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet      
    '�̼��� �߰�              
    Call InitComboBox
    Call InitVariables                      
    Call GetValue_ko441()
    Call SetDefaultVal
End Sub

<!--
'==========================================================================================
'   Event Name : txtAppFrDt    
'   Event Desc :
'==========================================================================================
-->
 Sub txtAppFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAppFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtAppFrDt.Focus
	End if
End Sub
<!--
'==========================================================================================
'   Event Name : txtAppToDt    
'   Event Desc :
'==========================================================================================
-->
 Sub txtAppToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtAppToDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtAppToDt.Focus
	End if
End Sub
<!--
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
-->
Sub txtAppFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

Sub txtAppToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery()
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
	
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

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    '�̼��� �߰� 
    Call InitComboBox
    Call InitData()
          
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor(-1, -1)
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_Curr,C_Cost,"C" ,"I","X","X")
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

<!--
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
-->
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
 
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)           
	Select Case Col
		Case C_CURR
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_Cost,"C" ,"X","X")
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_Curr,C_Cost,"C" ,"I","X","X")
		Case C_Cost
			Call FixDecimalPlaceByCurrency(frm1.vspdData,Row,C_Curr,C_Cost,"C" ,"X","X")
		Case C_PrcFlg	
			Call InitData()				
	End Select
End Sub

'========================================================================================================
'   Event Name : vspdData_EditMode
'   Event Desc : 
'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_Cost
            Call EditModeCheck(frm1.vspdData, Row, C_Curr, C_Cost,    "C" ,"I", Mode, "X", "X")
    End Select
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
   
	With frm1.vspdData 
 		ggoSpread.Source = frm1.vspdData
			 
		If Row > 0 And Col = C_OrderUnitPopUp Then
			.Col = Col
			.Row = Row
			Call OpenUnit(.text)
		Elseif Row > 0 And Col = C_PlantPopup Then
			.Col = C_PlantCd
			.Row = Row
			Call OpenSPlant(.text)
		Elseif Row > 0 And Col = C_ItemPopup Then
			.Col = C_ItemCd
			.Row = Row
			Call OpenSItem(.text)
		Elseif Row > 0 And Col = C_CurrPopup Then
			.Col = Col
			.Row = Row
			Call OpenCurr(.text)
		End if 
	End With
End Sub

<!--
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
-->
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then 
		If lgStrPrevKey <> "" Then       
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If 
			Call DisableToolBar(parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End if    
End Sub

<!--
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
-->
Function FncQuery() 
    Dim IntRetCD 
    Err.Clear                                               
    
    FncQuery = False                                        
    '-----------------------
    'Check previous data area
    '-----------------------
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
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
		If CompareDateByFormat(.txtAppFrDt.text,.txtAppToDt.text,.txtAppFrDt.Alt,.txtAppToDt.Alt, _
	              "970025",.txtAppFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtAppFrDt.text) <> "" And Trim(.txtAppToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","�ܰ�������", "X")
			Exit Function
		End if   
	End with
	'-----------------------
	'Query function call area
	'-----------------------
	If DbQuery = False Then Exit Function
	       
	Set gActiveElement = document.activeElement
	FncQuery = True   
End Function

<!--
'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
-->
Function FncNew() 
	Dim IntRetCD 
	    
	FncNew = False                                  
	Err.Clear                                       

	ggoSpread.Source = frm1.vspdData
	    
	'-----------------------
	'Check previous data area
	'-----------------------
	If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True  Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")          
	Call ggoOper.LockField(Document, "N")           
	Call SetDefaultVal
	Call InitVariables                              
	    
	Set gActiveElement = document.activeElement
	FncNew = True                                   
End Function

<!--
'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
-->
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                 
    
    Err.Clear                                       
    On Error Resume Next                           
    
    ggoSpread.Source = frm1.vspdData                

	If frm1.vspdData.MaxRows < 1 then
        Call DisplayMsgBox("17A002", "X","����", "X")
        Exit Function
    end if
    
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")  
        Exit Function
    End If
    
    If Not chkField(Document, "2")  Then              
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData                 
    If Not ggoSpread.SSDefaultCheck Then     
       Exit Function
    End If
 
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                       
End Function

<!--
'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
-->
Function FncCopy() 
	frm1.vspdData.ReDraw = False
	If frm1.vspdData.Maxrows < 1 then exit function
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.CopyRow
	SetSpreadColor1 frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow 
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Curr,C_Cost,"C" ,"I","X","X")
	Set gActiveElement = document.activeElement
	frm1.vspdData.ReDraw = True 
End Function

<!--
'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
-->
Function FncCancel() 
	If frm1.vspdData.Maxrows < 1 then exit function
	frm1.vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                               
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Curr,C_Cost,"C" ,"I","X","X")
	Set gActiveElement = document.activeElement
	frm1.vspdData.ReDraw = True
End Function

<!--
'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
-->
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow, iRow
	Dim lgF0
	Dim lgF1
	Dim lgF2
	Dim lgF3
	Dim lgF4
	Dim lgF5
	Dim lgF6
        
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End if
    End If
    
	With frm1
		If Not chkField(Document, "2") Then
			Exit Function
		End If

		CALL 	CommonQueryRs(" ORDER_UNIT_PUR ", " B_ITEM_BY_PLANT ", _
			" PLANT_CD = " & FilterVar(frm1.txtPlantCd1.Value, "''", "S") & " " , _
		    lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		    
		if lgF0 = "" or lgF0 = Null then
			
		Else
			lgF0 = split(lgF0,chr(11))
		End if
		
		.vspdData.ReDraw = False  
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		SetSpreadColor1 .vspdData.ActiveRow, .vspdData.ActiveRow + imRow -1
		
		'.vspdData.Row= .vspdData.ActiveRow
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow -1
			.vspdData.Row = iRow
			.vspdData.Col= C_OrderUnit
			.vspdData.Text = lgF0(0)

			If lgPLCd <> "" then 
				call .vspdData.SetText(C_PlantCd,iRow,lgPLCd)
			End If
			
			.vspdData.Row = iRow
			.vspdData.Col= C_AppDt
			.vspdData.Text = UNIFormatDate("<%=CDate(GetSvrDate)+1%>")
		Next
		.vspdData.ReDraw = True
	End With
    Set gActiveElement = document.ActiveElement   
    If Err.number = 0 Then 
		FncInsertRow = True                                                          '��: Processing is OK
	End If
End Function

<!--
'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
-->
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
    If frm1.vspdData.Maxrows < 1 then exit function
    With frm1.vspdData 
		.focus
		ggoSpread.Source = frm1.vspdData 
		    
		lDelRows = ggoSpread.DeleteRow
    End With    
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
-->
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
-->
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(parent.C_SINGLEMULTI)   
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
-->
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(parent.C_SINGLEMULTI , False)      
	Set gActiveElement = document.activeElement
End Function

<!--
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
-->
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

<!--
'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
-->
Function DbQuery() 
	Dim LngLastRow      
	Dim LngMaxRow       
	Dim LngRow          
	Dim strTemp         
	Dim StrNextKey      
	Dim strVal
		
	DbQuery = False
	    
	If LayerShowHide(1) = False then
		Exit Function 
	End if
	    
	Err.Clear                                                        
	    
	With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPlantCd1=" & Trim(.hdnPlant.value)   
			strVal = strVal & "&txtitemCd1=" & Trim(.hdnitem.value)
			strVal = strVal & "&txtItemGroupCd=" & Trim(.hdnitemGroup.value)
			strVal = strVal & "&txtAppFrDt=" & Trim(.hdnFrDt.value)
			strVal = strVal & "&txtAppToDt=" & Trim(.hdnToDt.value)
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPlantCd1=" & Trim(.txtPlantCd1.value)
			strVal = strVal & "&txtitemCd1=" & Trim(.txtitemCd1.value)
			strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
			strVal = strVal & "&txtAppFrDt=" & Trim(.txtAppFrDt.text)
			strVal = strVal & "&txtAppToDt=" & Trim(.txtAppToDt.text)
		End if
		
		.hdnmaxrow.value = .vspdData.MaxRows
		 
		Call RunMyBizASP(MyBizASP, strVal) 
	End With
	    
	DbQuery = True
End Function

'------------------------------------------  OpenItemGroup()  ----------------------------------------
' Name : OpenItemGroup()
' Description : HighItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�" 
	arrParam(1) = "B_ITEM_GROUP"    
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""   
	arrParam(5) = "ǰ��׷�"
 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
		   
	arrHeader(0) = "ǰ��׷�"
	arrHeader(1) = "ǰ��׷��"
	   
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
 
	IsOpenPop = False
 
	If arrRet(0) = "" Then
		frm1.txtItemGroupCd.focus
		Exit Function
	Else
		frm1.txtItemGroupCd.Value    = arrRet(0)   
		frm1.txtItemGroupNm.Value    = arrRet(1) 
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

<!--
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
-->
Function DbQueryOk()              
	Dim index
	Dim ii
	'-----------------------
	'Reset variables area
	'-----------------------
	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = parent.OPMD_UMODE           
	  
		Call ggoOper.LockField(Document, "Q")        
		Call SetToolBar("1110111100111111")
		Call SetSpreadLockAfterQuery()
	else
		Call ggoOper.LockField(Document, "N")
		Call SetToolBar("1110110100101111")
	End If

	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd1, "Q") 
	End If
	
	Call InitData()
 
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspddata.focus
	Else
		frm1.txtPlantCd1.focus 
	End If
	Set gActiveElement = document.activeElement
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

<!--
'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data save and display
'========================================================================================
-->
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal,strDel
	Dim ColSep, RowSep

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
    
	ColSep = parent.gColSep               
	RowSep = parent.gRowSep               

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ֱ� ������ ����[����,�ű�]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '�ֱ� ������ ����[����,�ű�]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

	With frm1
  
	If lgIntFlgMode = parent.OPMD_CMODE Then
	 .txtMode.value = parent.UID_M0002  '��� 
	Else
	 .txtMode.value = parent.UID_M0005  '���� 
	End If 
    
	'-----------------------
	'Data manipulate area
	'-----------------------
	lGrpCnt = 1
	strVal = ""
	strDel = ""

    For lRow = 1 To .vspdData.MaxRows

        Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
            Case ggoSpread.InsertFlag      
				strVal = "C" & ColSep    
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_Curr,lRow,"X","X")) & ColSep
				strVal = strVal & UNIConvDate(GetSpreadText(.vspdData,C_AppDt,lRow,"X","X")) & ColSep
				strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"),0) & ColSep 
				'�̼����߰� 
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PrcFlg,lRow,"X","X")) & ColSep 
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remark,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X")) & ColSep & lRow & RowSep
				lGrpCnt = lGrpCnt + 1
            Case ggoSpread.UpdateFlag      
				strVal = "U" & ColSep    
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_Curr,lRow,"X","X")) & ColSep
				strVal = strVal & UNIConvDate(GetSpreadText(.vspdData,C_AppDt,lRow,"X","X")) & ColSep
				strVal = strVal & UNIConvNum(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"),0) & ColSep 
				'�̼����߰� 
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PrcFlg,lRow,"X","X")) & ColSep 
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remark,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
				strVal = strVal & Trim(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X")) & ColSep & lRow & RowSep
                lGrpCnt = lGrpCnt + 1
            Case ggoSpread.DeleteFlag      
			    strDel = "D" & ColSep    
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X")) & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_Curr,lRow,"X","X")) & ColSep
				strDel = strDel & UNIConvDate(GetSpreadText(.vspdData,C_AppDt,lRow,"X","X")) & ColSep
				strDel = strDel & UNIConvNum(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"),0) & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_PrcFlg,lRow,"X","X")) & ColSep 
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_Remark,lRow,"X","X")) & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
				strDel = strDel & UNIConvNum(GetSpreadText(.vspdData,C_ItemCd,lRow,"X","X"),0) & ColSep & lRow & RowSep
                lGrpCnt = lGrpCnt + 1
        End Select
 
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
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
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		          
		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0 
		         End If
		       
		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If   
		         
		         iTmpDBuffer(iTmpDBufferCount) =  strDel         
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
		End Select   
   Next
 
	If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	If lGrpCnt > 1 Then
 		If LayerShowHide(1) = False Then
		   Exit Function 
		End If
	   
	   strVal = BIZ_PGM_ID & "?txtPlantCd1=" & frm1.txtPlantCd1.value 
	   strVal = strVal & "&txtItemGroupCd=" & frm1.txtItemGroupCd.value 
	   strVal = strVal & "&txtitemcd1=" & frm1.txtitemcd1.value 
	   strVal = strVal & "&txtAppFrDt=" & frm1.txtAppFrDt.value
	   strVal = strVal & "&txtAppToDt=" & frm1.txtAppToDt.value  
	   
		Call ExecMyBizASP(frm1, strVal)     
	End If
 
	End With
 
    DbSave = True                                           
End Function

<!--
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
-->
Function DbSaveOk()           
	Dim RowIndex

	Call InitVariables
	lgBlnFlgChgValue = false
	Call MainQuery()	
	frm1.vspdData.focus
End Function



'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
</HEAD>
<!-- '#########################################################################################################
'            6. Tag�� 
' ���: Tag�κ� ���� 
 ' 
 ' �ʵ��� ��� MaxLength=? �� ��� 
 ' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
  ' Normal Field�϶��� ������� ���� 
  ' Required Field�϶��� required�� �߰��Ͻʽÿ�.
  ' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
   ' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
 ' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
 ' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
 ' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### !-->
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�񺰴ܰ�</font></td>
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
         <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant('Plant1')" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
                  <INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm1" SIZE=20 MAXLENGTH=20 tag="14X">
                  
         <TD CLASS="TD5" NOWRAP>ǰ��</TD>
		<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemCd1" SIZE=10 MAXLENGTH=18 STYLE="text-transform:uppercase" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
							   <INPUT TYPE=TEXT ALT="ǰ��" NAME="txtItemNm1" SIZE=20 CLASS=protected readonly=true tag="14X" tabindex = -1></TD>
        </TR>
        <TR>
         <TD CLASS="TD5" NOWRAP>�ܰ�������</TD>
         <TD CLASS="TD6" NOWRAP>
          <table cellpadding=0 cellspacing=0>
           <tr>
            <td>
             <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ܰ�������" NAME="txtAppFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
            </td>
            <td>~<td>
            <td>
             <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="�ܰ�������" NAME="txtAppToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
            </td>
           </tr>
          </table>
         </TD>
         <TD CLASS="TD5" NOWRAP>ǰ��׷�</TD>
         <TD CLASS="TD6" NOWRAP>
				<INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemGroup()">
				<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 CLASS=protected readonly=true tag="1X" TABINDEX="-1">
         </TD>
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
        <TD HEIGHT="100%" WIDTH=100% COLSPAN=4>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
    
 <TR>
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
  </TD>
 </TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnItemGroup" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24" tabindex = -1>
<INPUT TYPE=HIDDEN NAME="hdnmaxrow"  tag="14">

	   
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>