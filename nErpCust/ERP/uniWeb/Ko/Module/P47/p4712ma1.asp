
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : Adjust Resource Compution
'*  3. Program ID           : p4712ma1
'*  4. Program Name         : Adjust Resource Compution
'*  5. Program Desc         : Adjust Resource Compution
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/30
'*  8. Modified date(Last)  : 2002/07/18
'*  9. Modifier (First)     : Jeon, JaeHyun
'* 10. Modifier (Last)      : Kang Seong Moon
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs">> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_LOOKUP_ID	= "p4712mb0.asp"								' Head�� ���������� ������ ���� ���� 

'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	= "p4712mb1.asp"								'��: BOR(Bill Of Resource) ���� ��ȸ 

'Grid 2 - Component Allocation
Const BIZ_PGM_QRY2_ID	= "p4712mb2.asp"								'��: Main Query(�ڿ���������) �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID	= "p4712mb3.asp"                                '��: �ڿ��Һ���� ó�� 
Const BIZ_PGM_LOOKUPRC_ID	= "p4712mb4.asp"                            '��: �ڿ����� LOOKUP ó�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

' Grid 1(vspdData1) - Operation 
Dim C_ResourceCd			'= 1
Dim C_ResourcePopup			'= 2
Dim C_ResourceNm			'= 3
Dim C_ConsumedDt			'= 4
Dim C_ConsumedQty			'= 5
Dim C_ResourceTypeNm		'= 6
Dim C_ResourceGroupCd		'= 7
Dim C_ResourceGroupNm		'= 8


' Grid 2(vspdData2) - Operation
Dim C_ResourceCd2			'= 1
Dim C_ResourceNm2			'= 2
Dim C_ResourceTypeNm2		'= 3
Dim C_ResourceGroupCd2		'= 4
Dim C_ResourceGroupNm2		'= 5
Dim	C_Rank2					'= 6
Dim	C_BOR_Efficiency2		'= 7
Dim C_ValidFromDt			'= 8
Dim C_ValidToDt				'= 9

Dim strDate
Dim BaseDate
Dim strYear
Dim strMonth
Dim strDay

BaseDate = "<%=GetSvrDate%>"

Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)			'��: �ʱ�ȭ�鿡 �ѷ����� ��¥ 

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount								'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgStrPrevKey

Dim lgLngCurRows
Dim lgCurrRow
Dim lgFlgQueryCnt
Dim lgSortKey
Dim lgSortKey2
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgRow         
'++++++++++++++++  Insert Your Code for Global Variables Assign  +++++++++++++++++++++++++++++++++++++

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
    
    lgStrPrevKey = ""							'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgRow = 0
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

End Sub

'========================================  2.2.1 SetCookieVal()  ======================================
'	Name : SetCookieVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=================================================================================================== 
Sub SetCookieVal()
	
	frm1.txtPlantCd.Value	= ReadCookie("txtPlantCd")
	frm1.txtPlantNm.value	= ReadCookie("txtPlantNm")
	frm1.txtItemCd.Value	= ReadCookie("txtItemCd")

	frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement

		
	frm1.txtProdOrderNo.Value	= ReadCookie("txtProdOrderNo")
	frm1.txtOprCd.Value			= ReadCookie("txtOprNo")
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtProdOrderNo", ""
	WriteCookie "txtOprNo", ""
		
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdNo)
	Call InitSpreadPosVariables(pvSpdNo)
	
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
		
	if pvSpdNo = "*" or pvSpdNo = "A" then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = C_ResourceGroupNm +1													'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0
		Call GetSpreadColumnPos("A")
		ggoSpread.SSSetEdit		C_ResourceCd,		"�ڿ��ڵ�",	10,,,10,2
		ggoSpread.SSSetButton 	C_ResourcePopup
		ggoSpread.SSSetEdit		C_ResourceNm,		"�ڿ���",	20
		ggoSpread.SSSetDate		C_ConsumedDt,		"�ڿ��Һ���",	13,	2,	parent.gDateFormat
		ggoSpread.SSSetTime		C_ConsumedQty,		"�ڿ��Һ�ð�",	13,2 ,1 ,1
		ggoSpread.SSSetEdit		C_ResourceTypeNm,	"�ڿ�����",	10
		ggoSpread.SSSetEdit		C_ResourceGroupCd,	"�ڿ��׷�",	10
		ggoSpread.SSSetEdit		C_ResourceGroupNm,	"�ڿ��׷��",	20
		
		Call ggoSpread.MakePairsColumn(C_ResourceCd,C_ResourcePopup)
		Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
		ggoSpread.SSSetSplit2(4)
		.ReDraw = true
		End With
	end if
	
	if pvSpdNo = "*" or pvSpdNo = "B" then
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = C_ValidToDt +1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxRows = 0
		Call GetSpreadColumnPos("B")
		ggoSpread.SSSetEdit		C_ResourceCd2,		"�ڿ��ڵ�", 10
		ggoSpread.SSSetEdit		C_ResourceNm2,		"�ڿ���",	20
		ggoSpread.SSSetEdit		C_ResourceTypeNm2,	"�ڿ�����", 10
		ggoSpread.SSSetEdit		C_ResourceGroupCd2, "�ڿ��׷�", 10
		ggoSpread.SSSetEdit		C_ResourceGroupNm2, "�ڿ��׷��", 20
		ggoSpread.SSSetFloat	C_Rank2,			"����",		10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BOR_Efficiency2,	"ȿ��",		10, "7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetDate		C_ValidFromDt,		"������",	11,	2,	parent.gDateFormat
		ggoSpread.SSSetDate		C_ValidToDt,		"������",	11,	2,	parent.gDateFormat
		'Call ggoSpread.MakePairsColumn(,)
		Call ggoSpread.SSSetColHidden(.MaxCols ,.MaxCols , True)
		ggoSpread.SSSetSplit2(2)	
		.ReDraw = true
		End With
    end if
	
	Call SetSpreadLock 
    
End Sub


'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadLock()

    With frm1
    '--------------------------------
    'Grid 1
    '--------------------------------
    ggoSpread.Source = frm1.vspdData1
	.vspdData1.ReDraw = False
	ggoSpread.SpreadLock C_ResourceCd,-1,C_ResourceCd
	ggoSpread.SpreadLock C_ResourcePopup,-1,C_ResourcePopup
	ggoSpread.SpreadLock C_ResourceNm,-1,C_ResourceNm
    ggoSpread.SpreadLock C_ConsumedDt,	-1
	ggoSpread.SpreadLock C_ResourceTypeNm, -1
	ggoSpread.SpreadLock C_ResourceGroupCd, -1
	ggoSpread.SpreadLock C_ResourceGroupNm, -1
	ggoSpread.SpreadLock frm1.vspdData1.MaxCols, -1, frm1.vspdData1.MaxCols
	
	ggoSpread.SpreadUnLock  C_ConsumedQty, -1 ,C_ConsumedQty
	ggoSpread.SSSetRequired  C_ConsumedQty, -1 , C_ConsumedQty
	.vspdData1.ReDraw = True
    
    '--------------------------------
    'Grid 2
    '--------------------------------
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()

    End With

End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1.vspdData1
    
    .Redraw = False
    
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.SSSetRequired  C_ResourceCd,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceNm,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	 C_ConsumedDt,			pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	 C_ConsumedQty,			pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceTypeNm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceGroupCd,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected C_ResourceGroupNm,		pvStartRow, pvEndRow

    .Redraw = True
    
    End With

End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
'Sub InitComboBox()

'End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	if pvSpdNo = "*" or pvSpdNo = "A" then
		'Grid 1(vspdData1) - Operation 
		C_ResourceCd			= 1
		C_ResourcePopup			= 2
		C_ResourceNm			= 3
		C_ConsumedDt			= 4
		C_ConsumedQty			= 5
		C_ResourceTypeNm		= 6
		C_ResourceGroupCd		= 7
		C_ResourceGroupNm		= 8
	end if

	if pvSpdNo = "*" or pvSpdNo = "B" then
		'Grid 2(vspdData2) - Operation
		C_ResourceCd2			= 1
		C_ResourceNm2			= 2
		C_ResourceTypeNm2		= 3
		C_ResourceGroupCd2		= 4
		C_ResourceGroupNm2		= 5
		C_Rank2					= 6
		C_BOR_Efficiency2		= 7
		C_ValidFromDt			= 8
		C_ValidToDt				= 9
	end if

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
		C_ResourceCd			= iCurColumnPos(1)
		C_ResourcePopup			= iCurColumnPos(2)
		C_ResourceNm			= iCurColumnPos(3)
		C_ConsumedDt			= iCurColumnPos(4)
		C_ConsumedQty			= iCurColumnPos(5)
		C_ResourceTypeNm		= iCurColumnPos(6)
		C_ResourceGroupCd		= iCurColumnPos(7)
		C_ResourceGroupNm		= iCurColumnPos(8)
	Case "B"
 		ggoSpread.Source = frm1.vspdData2
  		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_ResourceCd2			= iCurColumnPos(1)
		C_ResourceNm2			= iCurColumnPos(2)
		C_ResourceTypeNm2		= iCurColumnPos(3)
		C_ResourceGroupCd2		= iCurColumnPos(4)
		C_ResourceGroupNm2		= iCurColumnPos(5)
		C_Rank2					= iCurColumnPos(6)
		C_BOR_Efficiency2		= iCurColumnPos(7)
		C_ValidFromDt			= iCurColumnPos(8)
		C_ValidToDt				= iCurColumnPos(9) 
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
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
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

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
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
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
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
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtProdOrderNo.value = "" Then
		Call DisplayMsgBox("971012","X", "����������ȣ","X")
		frm1.txtProdOrderNo.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdOrderNo.value
	arrParam(2) = "Y"
	
	iCalledAspName = AskPRAspName("p4112pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4112pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetOprCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprCd.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource(ByVal strCode,ByVal strName,ByVal Row)

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
	arrParam(5) = "�ڿ�"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ�"		
    arrHeader(1) = "�ڿ���"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetResource(arrRet)
	End If
		
End Function

'------------------------------------------  OpenPartRef()  -------------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
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
   	With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With

	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4311ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4311ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -------------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()

	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
   	With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4111ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
End Function

'------------------------------------------  OpenProdRef()  -------------------------------------------------
'	Name : OpenProdRef()
'	Description : Production Reference
'---------------------------------------------------------------------------------------------------------
Function OpenProdRef()

	Dim arrRet
	Dim arrParam(2)
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

	arrParam(0) = Trim(frm1.txtPlantCd.value)	'��: ��ȸ ���� ����Ÿ 

   	With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
		arrParam(2) = .hOprNo.Value
	End With
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4411ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4411ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  -------------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Issue Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()

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
    With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4511ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4511ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P45/p4511ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenConsumRef()  -------------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
'---------------------------------------------------------------------------------------------------------
Function OpenConsumRef()

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
    With frm1
		If .hProdOrderNo.Value = "" Then Exit Function
		arrParam(1) = .hProdOrderNo.Value
	End With
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4412ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4412ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	'arrRet = window.showModalDialog("../P44/p4412ra1.asp", Array(arrParam(0), arrParam(1)), _
	'	"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetProdOrderNo()  --------------------------------------------------
'	Name : SetOprCd()
'	Description : Production Order Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprCd(byval arrRet)
	frm1.txtOprCd.Value    = arrRet(0)		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	With frm1.vspdData1
		ggoSpread.Source = frm1.vspdData1
			.Row = frm1.vspdData1.ActiveRow
			.Col = C_ResourceCd
			.Text = arrRet(0)
			
			'.Col = C_ResourceNm
			'.Text = arrRet(1)
			LookUpRc arrRet(0), .Row
	End with		
End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 


Sub InitData(ByVal lngStartRow)

End Sub

'==============================================================================
' Function : ConvToSec()
' Description : ����ÿ� �� �ð� �����͵��� �ʷ� ȯ�� 
'==============================================================================
Function ConvToSec(ByVal Str)

	If Str = "" Then
		ConvToSec = 0
	ElseIf Len(Trim(Str)) = 8 Then
		ConvToSec = CInt(Trim(Mid(Str,1,2))) * 3600 + CInt(Trim(Mid(Str,4,2))) * 60 + CInt(Trim(Mid(Str,7,2)))
	Else
		ConvToSec = -999999
	End If

End Function

'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Row = " & lRow & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Col = " & lCol & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Action = 0" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.focus" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Row = " & lRow & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Col = " & lCol & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.Action = 0" & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "frm1.vspdData1.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

'==========================================  2.5.2 LookupRc()  =======================================
'	Name : LookUpRc()
'	Description : Lookup WorkCenter using Keyboard
'===================================================================================================== 
Sub LookUpRc(ByVal strResourceCd, ByVal Row)
	Dim strVal		
	Call LayerShowHide(1)

	strVal = BIZ_PGM_LOOKUPRc_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtResourceCd=" & Trim(strResourceCd)								'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&Row=" & Row											'��: ��ȸ ���� ����Ÿ 
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
End Sub

Function LookUpRcOk(ByVal Rcd, ByVal RcdNm,ByVal RcdType, ByVal Rgcd, Byval RgcdNm, Byval Row)

	With frm1.vspdData1
				
		ggoSpread.Source = frm1.vspdData1
				
		.Row = Row
		.Col = C_ResourceNm
		.Text = RcdNm
		
		.Col = C_ResourceTypeNm
		.Text = RcdType
			
		.Col = C_ResourceGroupCd
		.Text = Rgcd
		
		.Col = C_ResourceGroupNm
		.Text = RgcdNm	
				
	End With

	IsOpenPop = False
		
End Function

Function LookUpRcNotOk(Byval Row)

	With frm1.vspdData1
				
		ggoSpread.Source = frm1.vspdData1
		
		.Row = Row
		.Col = C_ResourceCd
		.Text = ""
		.Col = C_ResourceNm
		.Text = ""
		.Col = C_ResourceTypeNm
		.Text = ""
		.Col = C_ResourceGroupCd
		.Text = ""
		.Col = C_ResourceGroupNm
		.Text = ""
	End With

	IsOpenPop = False
		
End Function


'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format

	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)        
	    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
    
	If ReadCookie("txtPlantCd") <> "" Then
		Call SetCookieVal
		frm1.txtOprCd.focus 
		Set gActiveElement = document.activeElement 
	Else	
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = UCase(parent.gPlant)
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtProdOrderNo.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================

'========================================================================================
' Function Name : vspdDat1a_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
  	
  	If lgIntFlgMode = Parent.OPMD_UMODE Then
  		Call SetPopupMenuItemInf("1101111111")         'ȭ�麰 ���� 
  	Else
  		Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
  	End If
  	
  	gMouseClickStatus = "SPC"   
     
  	Set gActiveSpdSheet = frm1.vspdData1
     
  	If frm1.vspdData1.MaxRows = 0 Then
  		Exit Sub
  	End If
  	
  	If Row <= 0 Then
  		ggoSpread.Source = frm1.vspdData1 
  		If lgSortKey = 1 Then
  			ggoSpread.SSSort Col					'Sort in Ascending
  			lgSortKey = 2
  		Else
  			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
  			lgSortKey = 1
  		End If
 	Else
  		'------ Developer Coding part (Start)
 	 	'------ Developer Coding part (End)
 	
  	End If

End Sub

'========================================================================================
' Function Name : vspdData2_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
  	gMouseClickStatus = "SP2C"   
     
  	Set gActiveSpdSheet = frm1.vspdData2
    
    Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
     
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
 
'=======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData1_Change(ByVal Col, ByVal Row)

	Dim strItemCd
	Dim strHndItemCd, strHndOprNo
	Dim i
	Dim strReqDt, strEndDt
	Dim	DblRqrdQty
	Dim strResourceCd
	
	ggoSpread.Source = frm1.vspdData1
	
	If Row < 1 Then Exit Sub
	
	With frm1.vspdData1	
	.Row = Row

	Select Case Col
			
			Case C_ResourceCd
				ggoSpread.Source = frm1.vspdData1
				frm1.vspdData1.Col = Col	
				frm1.vspdData1.Row = Row
				strResourceCd = frm1.vspdData1.text
					If frm1.vspdData1.Text <> "" Then
						Call LookUpRc(strResourceCd, Row)
					End If

				IsOpenPop = True  '?
			

		    Case C_ConsumedQty
						
			    ggoSpread.Source = frm1.vspdData1
				ggoSpread.UpdateRow Row
			
		    Case C_ConsumedDt
		    
				ggoSpread.Source = frm1.vspdData1
				ggoSpread.UpdateRow Row

		End Select

	End With

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
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
' 	If NewCol = C_XXX or Col = C_XXX Then
 '		Cancel = True
 '		Exit Sub
 '	End If
     ggoSpread.Source = frm1.vspdData1
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("A")
End Sub 

Sub vspdData2_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
 
 	 ggoSpread.Source = frm1.vspdData2
     Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
     Call GetSpreadColumnPos("B")
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
    dim pvSpdNo
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    pvSpdNo = gActiveSpdSheet.id
    Call InitSpreadSheet(pvSpdNo)
    Select Case pvSpdNo
	case "A"
		ggoSpread.Source = frm1.vspdData1
	case "B"
		ggoSpread.Source = frm1.vspdData2
	End Select 
	Call ggoSpread.ReOrderingSpreadData()
	 
End Sub 

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

Dim strCode
Dim strName

    With frm1.vspdData1
    
		ggoSpread.Source = frm1.vspdData1
		If Row < 1 Then Exit Sub

		Select Case Col

		    Case C_ResourcePopup
				.Col = C_ResourceCd
				.Row = Row
				strCode = .Text
				.Col = C_ResourceNm
				.Row = Row
				strName = .Text
				Call OpenResource(strCode, strName, Row)
				
				Call SetActiveCell(frm1.vspdData1,C_ResourceCd,Row,"M","X","X")
				Set gActiveElement = document.activeElement

		End Select

	End With

End Sub


'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If 
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1 ,NewTop) Then

		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
 			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub


Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If 
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then

		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
 			If DbDtlQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'========================================================================================
' Function Name : vspdData1_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SPC" Then
        gMouseClickStatus = "SPCR"
     End If
End Sub 
  
'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'		========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
     If Button = 2 And gMouseClickStatus = "SP2C" Then
        gMouseClickStatus = "SP2CR"
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
    
    FncQuery = False											'��: Processing is NG
    
    Err.Clear													'��: Protect system from crashing

    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")	'��: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
        
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
              
    Call InitVariables
    frm1.vspdData1.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
	lgFlgQueryCnt = 0

    '-----------------------
    'Query function call area
    '-----------------------
    
    If DbQuery = False Then
		 Call RestoreToolBar()
		 Exit Function												'��: Query db data
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
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
	
    Dim IntRetCD 
    Dim	LngRows
    
    FncSave = False												'��: Processing is NG
    
    Err.Clear                                                   '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")		'��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck Then						'��: Check required field(Multi area)
       Exit Function
    End If
    
    With frm1.vspdData1
     
    For LngRows = 1 To .MaxRows
      .Row = LngRows
      .Col = 0
      If .Text = ggoSpread.InsertFlag Or .Text = ggoSpread.UpdateFlag Or .Text = ggoSpread.DeleteFlag Then
		
		.Col = C_ConsumedQty					
			If .Text = "<%=ConvToTimeFormat(0)%>" Then
				Call DisplayMsgBox("189715", "x", "x", "x")
				Exit Function
			End If 
			
		.Col = C_ConsumedDt
			If CompareDateByFormat(.text,"<%=strDate%>","�ڿ��Һ���","������","970025",parent.gDateFormat,parent.gComDateType,True) = False Then
			  Exit Function               
			End If
	   End If
    Next
    
    End With
        
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function													'��: Save db data
    
    FncSave = True												'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
        
	If frm1.vspdData1.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData1.focus
    Set gActiveElement = document.activeElement 
	frm1.vspdData1.EditMode = True
	frm1.vspdData1.ReDraw = False
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.CopyRow
    frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
    frm1.vspdData1.Col = C_ResourceCd
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceNm
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceTypeNm
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceGroupCd
    frm1.vspdData1.Text = ""
    frm1.vspdData1.Col = C_ResourceGroupNm
    frm1.vspdData1.Text = ""
	SetSpreadColor frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow
    frm1.vspdData1.ReDraw = True
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 

  If frm1.vspdData1.MaxRows < 1 Then Exit Function	
  ggoSpread.EditUndo                                                  '��: Protect system from crashing


End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(ByVal pvRowCnt) 

    Dim IntRetCD
 	Dim imRow
 	Dim i
 	
 	On Error Resume Next
 	FncInsertRow = False

 	If IsNumeric(Trim(pvRowCnt)) Then
 		imRow = Cint(pvRowCnt)
 	Else
 		imRow = AskSpdSheetAddRowCount()
 		If imRow = "" Then
 			Exit Function
 		End If
 	End If
    
    With frm1
    	.vspdData1.focus
		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData1
			.vspdData1.ReDraw = False
	    ' 	ggoSpread.InsertRow
			ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
	     	SetSpreadColor .vspdData1.ActiveRow, .vspdData1.ActiveRow + imRow -1
			for i = 0 to imRow -1
				.vspdData1.Row = .vspdData1.ActiveRow + i
				.vspdData1.Col = C_ConsumedDt
				.vspdData1.text = strDate
				.vspdData1.Col = C_ConsumedQty
				.vspdData1.text = "<%=ConvToTimeFormat(0)%>"
				.vspdData1.ReDraw = True
			next
		'SetSpreadColor .vspdData1.ActiveRow		
	End With
    If Err.number = 0 Then FncInsertRow = True

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 

    Dim lDelRows
    Dim iDelRowCnt, i

    With frm1

		'.vspdData1.Row = frm1.vspdData1.ActiveRow
		'.vspdData1.Col = C_OrderStatus2
    	'If .vspdData1.Text = "ST" or .vspdData1.Text = "CL" Then
		'	Call DisplayMsgBox("189520", "x", "x", "x")
	    '		Exit Function
		'End IF    
		If .vspdData1.MaxRows < 1 Then Exit Function

		'Call DeleteMarkingHSheet()

    End With

	ggoSpread.Source = frm1.vspdData1
    lDelRows = ggoSpread.DeleteRow
    lgLngCurRows = lDelRows + lgLngCurRows

	'CopyToHSheet frm1.vspdData1.ActiveRow

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.fncPrint()                                                   '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData1							'��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")	'��: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

    Dim strVal
    
    DbQuery = False
    
    lgFlgQueryCnt = lgFlgQueryCnt + 1
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing

    With frm1

	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal =  BIZ_PGM_LOOKUP_ID	 & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey	    
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrdNo=" & Trim(.hProdOrderNo.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtOprNo=" & Trim(.hOprCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows		
	Else
		strVal =  BIZ_PGM_LOOKUP_ID	 & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.Value)			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtProdOrdNo=" & Trim(.txtProdOrderNo.Value)	'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtOprNo=" & Trim(.txtOprCd.Value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	lgOldRow = 1

	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
			Set gActiveElement = document.activeElement	
			If DbDtlQuery = False Then
				 Call RestoreToolBar()	
				 Exit Function 
			End If 
		End If
	End If

End Function


'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery�� ������ ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryNotOk()

'frm1.txtPlantCd.focus


End Function


'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOprCd
    
	boolExist = False
    With frm1
    
	DbDtlQuery = False   
    
		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.Value)
			strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.Value)
			strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
			strVal = strVal & "&txtRoutNo=" & Trim(.hRoutNo.Value)
			strVal = strVal & "&txtOprNo=" & Trim(.txtOprCd.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
			
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    End With

    DbDtlQuery = True

End Function



Function DbDtlQueryOk()												'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
	
	If lgFlgQueryCnt = 1 Then
		If lgIntFlgMode <> parent.OPMD_UMODE Then
			Call DbDtl2Query
		End If
	End If

End Function


'========================================================================================
' Function Name : DbDtl2Query
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtl2Query() 

Dim strVal
Dim boolExist
Dim lngRows
Dim strOprCd
    
	boolExist = False
    With frm1
    
	DbDtl2Query = False   
    
		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtOprNo=" & Trim(.hOprNo.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'��: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)		'��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtOprNo=" & Trim(.txtOprCd.Value)
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
			
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'��: �����Ͻ� ASP �� ���� 

    End With

    DbDtl2Query = True

End Function



Function DbDtl2queryOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
	
	Call SetToolbar("11001111001111")										'��: ��ư ���� ���� 
		
	frm1.vspdData1.ReDraw = False
   
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	lgAfterQryFlg = True
    Call InitData(LngMaxRow)
    frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1	
	frm1.vspdData1.ReDraw = True

End Function

Function DbDtl2querynotOk(ByVal LngMaxRow)												'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
	
	Call SetToolbar("11001101001111")										'��: ��ư ���� ���� 
		
	frm1.vspdData1.ReDraw = False
   
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	lgAfterQryFlg = True
    Call InitData(LngMaxRow)
    frm1.vspdData1.Col = 1
	frm1.vspdData1.Row = 1	
	frm1.vspdData1.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData()
    
End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================


Function DbSave() 

    Dim IntRows
    Dim strVal, strDel
	Dim ChkTimeVal
	Dim iColSep, iRowSep
    
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen						'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
	
	Dim iFormLimitByte						'102399byte
		
	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size

	Dim iTmpDBuffer							'������ ���� [����] 
	Dim iTmpDBufferCount					'������ ���� Position
	Dim iTmpDBufferMaxCount					'������ ���� Chunk Size
		
    DbSave = False                                                          '��: Processing is NG
      
    Call LayerShowHide(1)

    '-----------------------
    'Data manipulate area
    '-----------------------
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
	
	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
    
    '102399byte
    iFormLimitByte = parent.C_FORM_LIMIT_BYTE
    
    '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				

	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
	
	strCUTotalvalLen = 0 : strDTotalvalLen  = 0


	With frm1.vspdData1

		For IntRows = 1 To .MaxRows
    
			.Row = IntRows
			.Col = 0
	
			Select Case .Text
		    
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
		            
		            strVal = ""
		            
		            If .Text = ggoSpread.InsertFlag Then
						strVal = strVal & "C" & iColSep				'��: C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal = strVal & "U" & iColSep				'��: U=Update
					End If	
		            
					strVal = strVal & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					strVal = strVal & UCase(Trim(frm1.txtProdOrderNo.value)) & iColSep
					strVal = strVal & UCase(Trim(frm1.txtOprCd.value)) & iColSep

					.Col = C_ResourceCd
					strVal = strVal & Trim(.Text) & iColSep
					.Col = C_ConsumedDt		' ConsumedDt
					strVal = strVal & UNIConvDate(Trim(.Text)) & iColSep
					
					.Col = C_ConsumedQty	' ConsumedQty
					ChkTimeVal = ConvToSec(Trim(.Text))
					
					If ChkTimeVal = -999999	Then
						Call DisplayMsgBox("970029", vbInformation, "�ڿ��Һ�ð�", "", I_MKSCRIPT)
						Call SheetFocus(arrVal(1),8,I_MKSCRIPT)
						Response.End	
					Else
						strVal = strVal & ChkTimeVal & iRowSep
					End If

			    Case ggoSpread.DeleteFlag
			    
					strDel = "" 
					
					strDel = strDel & "D" & iColSep				'��: D=Delete
					strDel = strDel & UCase(Trim(frm1.txtPlantCd.value)) & iColSep
					strDel = strDel & UCase(Trim(frm1.txtProdOrderNo.value)) & iColSep
					strDel = strDel & UCase(Trim(frm1.txtOprCd.value)) & iColSep

					.Col = C_ResourceCd
					strDel = strDel & Trim(.Text) & iColSep
					.Col = C_ConsumedDt		' ConsumedDt
					strDel = strDel & UNIConvDate(Trim(.Text)) & iRowSep

			End Select
			
			.Col = 0
			
			Select Case .Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			    
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
			         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			         
				Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
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
	    
	End With
	
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
	
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)								'��: ���� �����Ͻ� ASP �� ���� 

    DbSave = True                                                           ' ��: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
    lgLngCurRows = 0                           'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0

	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
	
	Call RemovedivTextArea
	Call DbDtl2Query
	
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
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
' Function Name : RemovedivTextArea
' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

</SCRIPT>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : ConvToTimeFormat
' Description : �ð� �������� ���� 
'==============================================================================
Function ConvToTimeFormat(ByVal iVal)
	Dim iTime
	Dim iMin
	Dim iSec
				
	If IVal = 0 Then
		ConvToTimeFormat = "00:00:00"
	Else
		iMin = Fix(IVal Mod 3600)
		iTime = Fix(IVal /3600)
		
		iSec = Fix(iMin Mod 60)
		iMin = Fix(iMin / 60)
		
		ConvToTimeFormat = Right("0" & CStr(iTime),2) & ":" & Right("0" & CStr(iMin),2) & ":" & Right("0" & CStr(iSec),2)
		 
	End If
End Function

</script>
<!-- #Include file="../../inc/uni2kcm.inc" -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ��Һ���(������)</font></td> 
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>&nbsp;<A href="vbscript:OpenPartRef()">��ǰ����</A> | <A href="vbscript:OpenOprRef()">��������</A> | <A href="vbscript:OpenProdRef()">��������</A> | <A href="vbscript:OpenRcptRef()">�԰���</A> | <A href="vbscript:OpenConsumRef()">��ǰ�Һ񳻿�</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
								    <TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNo()"></TD>	
									<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprCd" SIZE=8 MAXLENGTH=3 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprCd()"></TD>
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
								<TD CLASS=TD5 NOWRAP>ǰ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="24" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4712ma1_I549825119_txtOrderQty.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�۾���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="24" ALT="�۾���">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" SIZE=20 tag="24"></TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME=txtOrderUnit SIZE=10 MAXLENGTH=10 tag="24xxxU" ALT="��������"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Tracking No</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="24" ALT="Tracking No">&nbsp;</TD>
								<TD CLASS=TD5 NOWRAP>��������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4712ma1_I176752752_txtProdQty.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo" SIZE=18 MAXLENGTH=20 tag="24" ALT="�����">&nbsp;</TD>
								<TD CLASS=TD5 NOWRAP>��ǰ����</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4712ma1_I264298713_txtGoodQty.js'></script>
								</TD>
							</TR>	
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p4712ma1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%" colspan=4>
									<script language =javascript src='./js/p4712ma1_B_vspdData2.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="txtRoutNo" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hOprNo" tag="24"><INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
<P ID="divTextArea"></P>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
