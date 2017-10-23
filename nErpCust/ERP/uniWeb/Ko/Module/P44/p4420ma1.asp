<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4420ma1.asp
'*  4. Program Name         : ��ǰ������� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : ADO : 189611saa
'*  7. Modified date(First) : 2002.11.21
'*  8. Modified date(Last)  : 2003/09/13
'*  9. Modifier (First)     : Park, Bumsoo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
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
Const BIZ_PGM_QRY_ID	= "p4420mb1.asp"							'��: �����Ͻ� ���� ASP�� 

'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================

' Grid 1(vspdData) - Operation
Dim C_ItemCd			
Dim C_ItemNm		
Dim C_ItemAcct		
Dim C_Spec	
Dim C_ProdQty		
Dim C_GoodQty		
Dim C_BadQty		
Dim C_RcptQty		
Dim C_Unit
Dim C_ItemGroupCd
Dim C_ItemGroupNm			

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							<%'Variable is for Dirty flag%>
Dim lgIntGrpCount								<%'Group View Size�� ������ ���� %>
Dim lgIntFlgMode								<%'Variable is for Operation Status%>

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey

Dim strDate
Dim iDBSYSDate

	iDBSYSDate = "<%=GetSvrDate%>"																'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	strDate = UniConvDateAToB(iDBSYSDate,parent.gServerDateFormat,parent.gDateFormat) 	'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 
         
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

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
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""							'initializes Previous Key 
    lgStrPrevKey2 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey = 1
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
	frm1.txtWorkDt.Text = UniConvDateAToB(strDate, Parent.gDateFormat, Parent.gDateFormatYYYYMM)
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","P","NOCOOKIE","MA") %>
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
		ggoSpread.Spreadinit "V20030913", ,Parent.gAllowDragDropSpread
				
		.ReDraw = false
				
		.MaxCols = C_ItemGroupNm + 1    
		.MaxRows = 0    
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit 	C_ItemCd,       "ǰ��"			, 18
		ggoSpread.SSSetEdit 	C_ItemNm,       "ǰ���"		, 25
		ggoSpread.SSSetEdit 	C_Spec,			"�԰�"		, 25           
		ggoSpread.SSSetEdit 	C_ItemAcct,     "ǰ�����"		, 10
		ggoSpread.SSSetFloat	C_ProdQty,      "�ѻ��귮"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_GoodQty,		"��ǰ����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_BadQty,		"�ҷ�����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_RcptQty,		"�԰����"		,15,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit 	C_Unit,			"����",8
		ggoSpread.SSSetEdit 	C_ItemGroupCd,	"ǰ��׷�",	15
		ggoSpread.SSSetEdit		C_ItemGroupNm,	"ǰ��׷��", 30
		
		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
		
		ggoSpread.SSSetSplit2(2)
		
		Call SetSpreadLock 
		
		.ReDraw = true    
    
    End With
	
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

	'--------------------------------
	'Grid 1
	'--------------------------------    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
        
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'========================== 2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================
Sub InitComboBox()

    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD < " & FilterVar("30", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
    
    frm1.cboItemAcct.value = ""

End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables()

	C_ItemCd			= 1
	C_ItemNm			= 2
	C_Spec				= 3
	C_ItemAcct			= 4
	C_ProdQty			= 5
	C_GoodQty			= 6
	C_BadQty			= 7 
	C_RcptQty			= 8
	C_Unit			    = 9
	C_ItemGroupCd		= 10
	C_ItemGroupNm		= 11


End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_ItemCd			= iCurColumnPos(1)
			C_ItemNm			= iCurColumnPos(2)
			C_Spec				= iCurColumnPos(3)
			C_ItemAcct			= iCurColumnPos(4)
			C_ProdQty			= iCurColumnPos(5)
			C_GoodQty			= iCurColumnPos(6)
			C_BadQty			= iCurColumnPos(7)
			C_RcptQty			= iCurColumnPos(8)
			C_Unit			    = iCurColumnPos(9)
			C_ItemGroupCd		= iCurColumnPos(10)
			C_ItemGroupNm		= iCurColumnPos(11)
			
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
'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenDayDetailRef()  -------------------------------------------------
'	Name : OpenDayDetailRef()
'	Description : OpenDayDetailRef Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenDayDetailRef()
	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	
	Dim strDate1
	Dim strYear1
	Dim strMonth1
	Dim strDay1

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4420RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4420RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	Call ExtractDateFrom(frm1.txtWorkDt.Text,frm1.txtWorkDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")
	
	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))	
	arrParam(1) = strDate1
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItemCd
	arrParam(2) = UCase(Trim(frm1.vspdData.Value))
	frm1.vspdData.Col = C_ItemNm
	arrParam(3) = frm1.vspdData.Value
	arrParam(4) = frm1.cboItemAcct.value
	arrParam(5) = UCase(Trim(frm1.txtWcCd.value))
	arrParam(6) = UCase(Trim(frm1.txtShiftCd.value))
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2),arrParam(3), arrParam(4), arrParam(5),arrParam(6)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenWcRef()  -------------------------------------------------
'	Name : OpenWcRef()
'	Description : OpenWcRef Reference PopUp
'----------------------------------------------------------------------------------------------------------
Function OpenWCRef()
	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	
	Dim strDate1
	Dim strYear1
	Dim strMonth1
	Dim strDay1

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4420RA2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4420RA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	Call ExtractDateFrom(frm1.txtWorkDt.Text,frm1.txtWorkDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))	
	arrParam(1) = strDate1
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItemCd
	arrParam(2) = UCase(Trim(frm1.vspdData.Value))
	frm1.vspdData.Col = C_ItemNm
	arrParam(3) = frm1.vspdData.Value
	arrParam(4) = frm1.cboItemAcct.value
	arrParam(5) = UCase(Trim(frm1.txtWcCd.value))
	arrParam(6) = UCase(Trim(frm1.txtShiftCd.value))
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2),arrParam(3), arrParam(4), arrParam(5),arrParam(6)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False
End Function    

'------------------------------------------  OpenShiftRef()  ---------------------------------------------
'	Name : OpenShiftRef()
'	Description : OpenShiftRef Reference PopUp
'--------------------------------------------------------------------------------------------------------- 

Function OpenShiftRef()
	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	
	Dim strDate1
	Dim strYear1
	Dim strMonth1
	Dim strDay1

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

    If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4420RA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4420RA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	Call ExtractDateFrom(frm1.txtWorkDt.Text,frm1.txtWorkDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))	
	arrParam(1) = strDate1
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItemCd
	arrParam(2) = UCase(Trim(frm1.vspdData.Value))
	frm1.vspdData.Col = C_ItemNm
	arrParam(3) = frm1.vspdData.Value
	arrParam(4) = frm1.cboItemAcct.value
	arrParam(5) = UCase(Trim(frm1.txtWcCd.value))
	arrParam(6) = UCase(Trim(frm1.txtShiftCd.value))
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2),arrParam(3), arrParam(4), arrParam(5),arrParam(6)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False
End Function    

'------------------------------------------  OpenOrderRef()  ---------------------------------------------
'	Name : OpenOrderRef()
'	Description : OpenOrderRef Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenOrderRef()
	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	
	Dim strDate1
	Dim strYear1
	Dim strMonth1
	Dim strDay1

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

     If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4420RA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4420RA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	Call ExtractDateFrom(frm1.txtWorkDt.Text,frm1.txtWorkDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))	
	arrParam(1) = strDate1
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItemCd
	arrParam(2) = UCase(Trim(frm1.vspdData.Value))
	frm1.vspdData.Col = C_ItemNm
	arrParam(3) = frm1.vspdData.Value
	arrParam(4) = frm1.cboItemAcct.value
	arrParam(5) = UCase(Trim(frm1.txtWcCd.value))
	arrParam(6) = UCase(Trim(frm1.txtShiftCd.value))
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2),arrParam(3), arrParam(4), arrParam(5),arrParam(6)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False
End Function

'------------------------------------------  OpenTrackingRef()  ---------------------------------------------
'	Name : OpenTrackingRef()
'	Description : OpenTrackingRef Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingRef()
	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	
	Dim strDate1
	Dim strYear1
	Dim strMonth1
	Dim strDay1

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

     If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4420RA5")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4420RA5", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	Call ExtractDateFrom(frm1.txtWorkDt.Text,frm1.txtWorkDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))	
	arrParam(1) = strDate1
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = C_ItemCd
	arrParam(2) = UCase(Trim(frm1.vspdData.Value))
	frm1.vspdData.Col = C_ItemNm
	arrParam(3) = frm1.vspdData.Value
	arrParam(4) = frm1.cboItemAcct.value
	arrParam(5) = UCase(Trim(frm1.txtWcCd.value))
	arrParam(6) = UCase(Trim(frm1.txtShiftCd.value))
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2),arrParam(3), arrParam(4), arrParam(5),arrParam(6)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: Yes; status: No;")

	IsOpenPop = False
End Function

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
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus	
	
End Function

'------------------------------------------  OpenItemCd()  -----------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	 If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(frm1.txtItemCd.value) 						
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus	
	
End Function

'------------------------------------------  OpenItemGroup()  -------------------------------------------------
'	Name : OpenItemGroup()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
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

'------------------------------------------  OpenConWC()  ------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	 If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		'Call displaymsgbox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWcCd.Value)									' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    arrField(2) = "INSIDE_FLG"
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    arrHeader(2) = "�۾���Ÿ��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConWC(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWcCd.focus	
	
End Function


'------------------------------------------  OpenShiftCd()  ----------------------------------------------
'	Name : OpenShiftCd()
'	Description : Shift Popup
'---------------------------------------------------------------------------------------------------------
Function OpenShiftCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	 If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Shift �˾�"											' �˾� ��Ī 
	arrParam(1) = "P_SHIFT_HEADER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtShiftCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")	' Where Condition
	arrParam(5) = "Shift"												' TextBox ��Ī 
	 
    arrField(0) = "SHIFT_CD"												' Field��(0)
    arrField(1) = "DESCRIPTION"												' Field��(1)
    
    arrHeader(0) = "Shift"												' Header��(0)
    arrHeader(1) = "Shift ��"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetShiftCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtShiftCd.focus	
	
End Function
'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetItemInfo()  -----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)		
	End With
End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemGroup()
'	Description : Item Group Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetConWC()  --------------------------------------------------
'	Name : SetConWC()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------- 
Function SetConWC(byval arrRet)
	frm1.txtWcCd.Value    = arrRet(0)		
	frm1.txtWcNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetShiftCd()  -------------------------------------------------
'	Name : SetShiftCd()
'	Description : Condition Shift Popup���� Return�Ǵ� �� setting
'-----------------------------------------------------------------------------------------------------------
Function SetShiftCd(byval arrRet)
	frm1.txtShiftCd.Value    = arrRet(0)			
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
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.FormatDate(frm1.txtWorkDt, parent.gDateFormat, "2")
  
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    
       '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal

    Call InitVariables                                                      '��: Initializes local global variables
  
    Call InitComboBox
    
    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
    
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

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtWorkDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtWorkDt.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtWorkDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtWorkDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtWorkDt_KeyDown(keycode, shift)
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
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
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
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True    
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

    End With

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

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If
	
	If frm1.txtWcCd.value = "" Then
		frm1.txtWcCd.value = ""
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")								'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables											'��: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function																'��: Query db data
	End If
	
    FncQuery = True																'��: Processing is OK
    
    
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
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
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
    Call parent.FncExport(parent.C_MULTI)                                          '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      '��: Protect system from crashing
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

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

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
    
    Err.Clear                                                               '��: Protect system from crashing

   With frm1
   	 
	
	Call ExtractDateFrom(.txtWorkDt.Text,.txtWorkDt.UserDefinedFormat,parent.gComDateType,strYear1,strMonth1,strDay1)  '��: Extract Date data
	strDate1 = UniConvYYYYMMDDToDate(parent.gServerDateFormat, strYear1, strMonth1, "01")
   
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & UCase(Trim(.hPlantCd.value))			'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & UCase(Trim(.hItemCd.value))			'��: ��ȸ ���� ����Ÿ		
		strVal = strVal & "&txtWorkDt=" & UCase(Trim(.hWorkDt.value))
		strVal = strVal & "&cboItemAcct=" & .hItemAcct.value
		strVal = strVal & "&txtWcCd=" & UCase(Trim(.hWcCd.value))
		strVal = strVal & "&txtShliftCd=" & UCase(Trim(.hShiftCd.value))
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)		
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
   Else
		strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & UCase(Trim(.txtPlantCd.value))		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtItemCd=" & UCase(Trim(.txtItemCd.value))			'��: ��ȸ ���� ����Ÿ		
		strVal = strVal & "&txtWorkDt=" & strDate1
		strVal = strVal & "&cboItemAcct=" & .cboItemAcct.value
		strVal = strVal & "&txtWcCd=" & UCase(Trim(.txtWcCd.value))
		strVal = strVal & "&txtShliftCd=" & UCase(Trim(.txtShiftCd.value))
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
   End If
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	Call SetToolBar("11000000000111")										'��: ��ư ���� ���� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	
	lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True	
	
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
    Call InitSpreadSheet
    ggoSpread.Source = gActiveSpdSheet
		
	Call ggoSpread.ReOrderingSpreadData()

End Sub 


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ�������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenDayDetailRef()">�Ϻ���</A>&nbsp;|&nbsp;<A href="vbscript:OpenWCRef()">�۾��庰</A>&nbsp;|&nbsp;<A href="vbscript:OpenShiftRef()">Shift ��</A>&nbsp;|&nbsp;<A href="vbscript:OpenOrderRef()">������</A>&nbsp;|&nbsp;<A href="vbscript:OpenTrackingRef()">Tracking No.��</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>�۾����</TD> 
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4420ma1_I414607453_txtWorkDt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14" ALT="ǰ���"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" ALT="ǰ�����" STYLE="Width: 160px;" tag="11"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="ǰ��׷��"></TD>
									<TD CLASS=TD5 NOWRAP>�۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtWcCd" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenConWC()"> <INPUT TYPE=TEXT NAME="txtWcNm" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>Shift</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd" SIZE=5 MAXLENGTH=2 tag="11xxxU" ALT="Shift"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShiftCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenShiftCd()"></TD>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="100%">
								<TD WIDTH="50%">
									<script language =javascript src='./js/p4420ma1_I888265126_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hShiftCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hWorkDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24"><INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
