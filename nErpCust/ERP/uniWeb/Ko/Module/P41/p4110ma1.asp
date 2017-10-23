
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4110ma1.asp
'*  4. Program Name         : Explosion Prod. Order
'*  5. Program Desc         : p4110mb1.asp p4110mb2.asp  
'*  6. Comproxy List        : ADO : 189702saa, 189702sab
'*  7. Modified date(First) : 2000/12/12
'*  8. Modified date(Last)  : 2002/08/20
'*  9. Modifier (First)     : Park, Bum Soo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>������������</TITLE>
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT> 
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit															'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Grid 1 - Operation
Const BIZ_PGM_QRY1_ID	 = "p4110mb1.asp"		
Const BIZ_PGM_RUN_ID	 = "p4110mb2.asp"
Const BIZ_PGM_CONFIRM_ID = "p4110mb3.asp"						'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMPTOORDER_ID1	= "p4111ma1"
Const BIZ_PGM_JUMPTOORDER_ID2	= "p4112ma1"

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const C_SHEETMAXROWS = 30

' Grid 1(vspdData1) - Operation 
Dim C_Select1
Dim C_ItemCd1		'= 1
Dim C_ItemNm1		'= 2
Dim C_Spec1			'
Dim C_StartDt1		'= 3
Dim C_DueDt1		'= 4
Dim C_PlanQty1		'= 5
Dim C_TrackingNo1	'= 6
Dim C_MpsNo1		'= 7
Dim C_SplitSeq1		'= 8
Dim C_BOMNo1		'= 9

' Grid 2(vspdData2) - Operation 
Dim C_Select2
Dim C_ItemCd2		'= 1
Dim C_ItemNm2		'= 2
Dim C_Spec2
Dim C_StartDt2		'= 3
Dim C_DueDt2		'= 4
Dim C_PlanQty2		'= 5
Dim C_TrackingNo2	'= 6
Dim C_MpsNo2		'= 7
Dim C_SplitSeq2		'= 8
Dim C_BOMNo2		'= 9
Dim C_ProcurType2	'= 10
Dim C_ProcurNm2		'= 11
Dim C_SelectForPurQty2 '= 12

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue							'Variable is for Dirty flag
Dim lgIntGrpCount							'Group View Size�� ������ ���� 
Dim lgIntFlgMode								'Variable is for Operation Status

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows

Dim lgSortKey1
Dim lgSortKey2

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lstrPgmID
Dim lgInvCloseDt
Dim lgDateCheckFlg
Dim lgButtonSelection
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
    lgSortKey1 = 1
    lgSortKey2 = 2
    lgButtonSelection = "DESELECT"
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "��ü����"
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
	Dim LocSvrDate
	Dim DtSvrDate
	Dim DtInvCloseDt
	Dim DtStartDt

	If Trim(ReadCookie("txtPlantCd")) <> "" Then
		frm1.txtPlantCd.Value		= ReadCookie("txtPlantCd")
		frm1.txtPlantNm.value		= ReadCookie("txtPlantNm")
		frm1.txtItemCd.Value		= ReadCookie("txtItemCd")
		frm1.txtItemNm.value		= ReadCookie("txtItemNm")
		frm1.txtSpecification.value	= ReadCookie("txtSpecification")
		frm1.txtProdOrderNo.Value	= ReadCookie("txtProdOrderNo")
		frm1.txtPlanOrderNo.value	= ReadCookie("txtPlanOrderNo")
		frm1.txtOrderQty.Value		= ReadCookie("txtOrderQty")
		frm1.txtOrderUnit.Value		= ReadCookie("txtOrderUnit")
		frm1.txtStartDt.Text		= ReadCookie("txtPlanStartDt")		
		frm1.txtEndDt.Text			= ReadCookie("txtPlanEndDt")
		lgInvCloseDt				= ReadCookie("txtInvCloseDt")
		lstrPgmID = ReadCookie("txtPGMID")
	End If		
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtItemCd", ""	
	WriteCookie "txtItemNm", ""	
	WriteCookie "txtSpecification", ""
	WriteCookie "txtProdOrderNo", ""
	WriteCookie "txtPlanOrderNo", ""
	WriteCookie "txtOrderQty", ""
	WriteCookie "txtOrderUnit", ""
	WriteCookie "txtPlanStartDt", ""
	WriteCookie "txtPlanEndDt", ""
	WriteCookie "txtInvCloseDt", ""
	WriteCookie "txtPGMID", ""

	LocSvrDate = "<%=GetSvrDate%>"
	DtSvrDate	 = UniConvDateAToB(LocSvrDate, parent.gDateFormat, parent.gServerDateFormat)
	DtInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)
	DtStartDt    = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	frm1.txtExecFromDt.text = UniConvDateAToB(LocSvrDate, parent.gServerDateFormat, parent.gDateFormat)	
	
	frm1.btnAutoSel.disabled = True
	frm1.btnAutoSel.value = "��ü����"
	
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'======================================================================================== 
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)   
	
	 If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1     
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20060108", , Parent.gAllowDragDropSpread
			.ReDraw = False
			
			.MaxCols = C_BOMNo1 + 1   
			.MaxRows = 0 
			
			Call GetSpreadColumnPos("A")
			
			ggoSpread.SSSetCheck	C_Select1,		 ""					,2,,,1
			ggoSpread.SSSetEdit 	C_ItemCd1,       "ǰ��"			,18
			ggoSpread.SSSetEdit 	C_ItemNm1,       "ǰ���"		,25           
			ggoSpread.SSSetEdit 	C_Spec1,		 "�԰�"			,25
			ggoSpread.SSSetDate 	C_StartDt1,		 "����������"	,10, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_DueDt1,		 "�ϷΌ����"	,10, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_PlanQty1,		 "��������"		,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_TrackingNo1,	 "Tracking No."	,25
			ggoSpread.SSSetEdit 	C_MpsNo1,		 "MPS No."	,8
			ggoSpread.SSSetEdit 	C_SplitSeq1,	 "����"	,8
			ggoSpread.SSSetEdit 	C_BOMNo1,		 "BOM Type"	,8
			
			'Call ggoSpread.MakePairsColumn(,)
			'Call ggoSpread.SSSetColHidden( C_MpsNo1, C_BOMNo1, True)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)							'frozen ����߰� 
			
			Call SetSpreadLock("A")
				
			.ReDraw = true
		End With
	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then	
		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
		With frm1.vspdData2
    
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20060110", , Parent.gAllowDragDropSpread
			
			.ReDraw = false
			
			.MaxCols = C_SelectForPurQty2 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ    
			.MaxRows = 0
			
			ggoSpread.SSSetCheck	C_Select2,		 ""					,2,,,1 	
			ggoSpread.SSSetEdit 	C_ItemCd2,       "ǰ��"			,18 
			ggoSpread.SSSetEdit 	C_ItemNm2,       "ǰ���"		,25           
			ggoSpread.SSSetEdit 	C_Spec2,		 "�԰�"			,25
			ggoSpread.SSSetDate 	C_StartDt2,		 "���ֿ�����"	,10, 2, parent.gDateFormat
			ggoSpread.SSSetDate 	C_DueDt2,		 "���⿹����"	,10, 2, parent.gDateFormat       
			ggoSpread.SSSetFloat	C_PlanQty2,		 "��������"		,15,parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_TrackingNo2,	 "Tracking No."	,25
			ggoSpread.SSSetEdit 	C_MpsNo2,		 "MPS No."	,8
			ggoSpread.SSSetEdit 	C_SplitSeq2,	 "����"	,8
			ggoSpread.SSSetEdit 	C_BOMNo2,		 "BOM Type"	,8
			ggoSpread.SSSetEdit		C_ProcurType2,	 "���ޱ���"	,8
			ggoSpread.SSSetEdit		C_ProcurNm2,	 "���ޱ���"	,12 
			ggoSpread.SSSetCheck	C_SelectForPurQty2,""					,2,,,1 
			
			
			'Call ggoSpread.MakePairsColumn(,)
			'Call ggoSpread.SSSetColHidden( C_MpsNo2, C_ProcurType2, True)
 			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)							'frozen ����߰� 
			
			Call SetSpreadLock("B")
			
			.ReDraw = true
    
		End With
		
	End If    
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadLock(ByVal pvSpdNo)

    With frm1
		If pvSpdNo = "A" Then
			'--------------------------------
			'Grid 1
			'--------------------------------
			ggoSpread.Source = .vspdData1
			ggoSpread.SpreadLock -1, -1	' Set Lock Property : Spread 1
			ggoSpread.spreadUnLock C_Select1, -1, C_Select1
			
		End If
	
		If pvSpdNo = "B" Then    
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.SpreadLock -1, -1	' Set Lock Property : Spread 1
			ggoSpread.spreadUnLock C_Select2, -1, C_Select2
		End If
    End With
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================== 
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'=========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Operation 
		C_Select1			= 1
		C_ItemCd1			= 2
		C_ItemNm1			= 3
		C_Spec1				= 4
		C_StartDt1			= 5
		C_DueDt1			= 6
		C_PlanQty1			= 7
		C_TrackingNo1		= 8
		C_MpsNo1			= 9
		C_SplitSeq1			= 10
		C_BOMNo1			= 11
	End If
		
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2) - Operation 
		C_Select2			= 1
		C_ItemCd2			= 2
		C_ItemNm2			= 3
		C_Spec2				= 4
		C_StartDt2			= 5
		C_DueDt2			= 6
		C_PlanQty2			= 7
		C_TrackingNo2		= 8
		C_MpsNo2			= 9
		C_SplitSeq2			= 10
		C_BOMNo2			= 11
		C_ProcurType2		= 12
		C_ProcurNm2			= 13
		C_SelectForPurQty2	= 14
	End If
		
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
 			
 			C_Select1			= iCurColumnPos(1)
			C_ItemCd1			= iCurColumnPos(2)
			C_ItemNm1			= iCurColumnPos(3)
			C_Spec1				= iCurColumnPos(4)
			C_StartDt1			= iCurColumnPos(5)
			C_DueDt1			= iCurColumnPos(6)
			C_PlanQty1			= iCurColumnPos(7)
			C_TrackingNo1		= iCurColumnPos(8)
			C_MpsNo1			= iCurColumnPos(9)
			C_SplitSeq1			= iCurColumnPos(10)
			C_BOMNo1			= iCurColumnPos(11)
		
		Case "B"
 			ggoSpread.Source = frm1.vspdData2 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 			
 			C_Select1			= iCurColumnPos(1)
 			C_ItemCd2			= iCurColumnPos(2)
			C_ItemNm2			= iCurColumnPos(3)
			C_Spec2				= iCurColumnPos(4)
			C_StartDt2			= iCurColumnPos(5)
			C_DueDt2			= iCurColumnPos(6)
			C_PlanQty2			= iCurColumnPos(7)
			C_TrackingNo2		= iCurColumnPos(8)
			C_MpsNo2			= iCurColumnPos(9)
			C_SplitSeq2			= iCurColumnPos(10)
			C_BOMNo2			= iCurColumnPos(11)
			C_ProcurType2		= iCurColumnPos(12)
			C_ProcurNm2			= iCurColumnPos(13)
			C_SelectForPurQty2	= iCurColumnPos(14)
 			
 	End Select
 
End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()

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
'------------------------------------------  OpenPartRef()  -------------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRef()
	Dim arrRet
	Dim arrParam(11)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("P4110RA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4110RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ	
	arrParam(1) = Trim(frm1.txtPlantNm.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = Trim(frm1.txtProdOrderNo.value)
	arrParam(4) = Trim(frm1.txtPlanOrderNo.value)
	arrParam(5) = Trim(frm1.txtOrderQty.Value)
	arrParam(6) = Trim(frm1.txtStartDt.Text)
	arrParam(7) = Trim(frm1.txtEndDt.Text)
	arrParam(8) = Trim(frm1.chkInvStock.checked)
	arrParam(9) = Trim(frm1.chkSFStock.checked)
	arrParam(10) = Trim(frm1.chkForward.checked)
	arrParam(11) = Trim(frm1.txtItemNm.value)
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3), arrParam(4), arrParam(5), arrParam(6), arrParam(7), arrParam(8), arrParam(9), arrParam(10),arrParam(11)), _
		"dialogWidth=960px; dialogHeight=420px; center: Yes; help: No; resizable: Yes; status: No; scrollbar: Yes")
	
	IsOpenPop = False

End Function


'------------------------------------------  OpenErrorList()  -------------------------------------------------
'	Name : OpenErrorList()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenErrorList()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value  = "" Then
		call DisplayMsgBox("220705", "X","X","X")
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtPlanOrderNo.value)
	
	iCalledAspName = AskPRAspName("p4110ra2")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4110ra2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName , Array(window.parent, arrParam(0), arrParam(1)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function


'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function Run()
	Err.Clear																'��: Protect system from crashing
    Run = False																'��: Processing is NG
	
	If lgDateCheckFlg = "False1" Then 
		Call DisplayMsgBox("189250", "x", "x", "x")
		Exit Function
'	ElseIf lgDateCheckFlg = "False2" Then 
'		Call DisplayMsgBox("189251", "x", "x", "x")
'		Exit Function
	End If
	
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Dim strVal
    Call LayerShowHide(1)
    
    If Not chkfield(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
    
    With frm1
	
		strVal = BIZ_PGM_RUN_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtPlanOrderNo=" & Trim(.txtPlanOrderNo.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtExecFromDt=" & Trim(.txtExecFromDt.text)
		strVal = strVal & "&txtStartDt=" & Trim(.txtStartDt.text)
		strVal = strVal & "&txtEndDt=" & Trim(.txtEndDt.text)
		strVal = strVal & "&chkInvStock=" & Trim(.chkInvStock.checked)
		strVal = strVal & "&chkSFStock=" & Trim(.chkSFStock.checked)
		strVal = strVal & "&chkForward=" & Trim(.chkForward.checked)		
		strVal = strVal & "&txtFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtUserId=" & parent.gUsrID
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    
    End With	
    
    Run = True 
    lgBlnFlgChgValue = False    
    
End Function    

Function Confirm()

	Dim sp1ChgFlg,sp2ChgFlg
	Dim IntRetCD

	Err.Clear																'��: Protect system from crashing
    Confirm = False														'��: Processing is NG
	
	
	ggoSpread.Source = frm1.vspdData1                          '��: Preset spreadsheet pointer 
	sp1ChgFlg = ggoSpread.SSCheckChange
	
	ggoSpread.Source = frm1.vspdData2                          '��: Preset spreadsheet pointer 
	sp2ChgFlg = ggoSpread.SSCheckChange
	
	
	If sp1ChgFlg = False And sp2ChgFlg = False Then
		IntRetCD = DisplayMsgBox("900001", "x", "x", "x")     '��: Display Message(There is no changed data.)
        Exit Function
	End If
	
	
	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "x", "x")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If Not chkfield(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
    
    Call ConfirmSave()
    
    Confirm = True 
    lgBlnFlgChgValue = False
End Function

Function ConfirmSave()

    Dim strVal
    Dim txtVal1, txtVal2
    Dim TmpArr1, TmpArr2
    Dim iColSep, iRowSep
    Dim IntQtyRow
    
    Dim pvCnt
    
	ConfirmSave = True
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = parent.UID_M0002												'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		.txtUserId.value = parent.gUsrID
		
		Redim TmpArr1(.vspdData1.MaxRows + .vspdData2.MaxRows)
		
		iColSep = parent.gColSep : iRowSep = parent.gRowSep  
		
		For pvCnt = 1 To .vspdData1.MaxRows
			.vspdData1.Row = pvCnt
			  .vspdData1.CoL = C_Select1
		    If  .vspdData1.Text = "1"   then
				txtVal1 = ""
				.vspdData1.Col = C_ItemCd1 : txtVal1 = txtVal1 & Trim(.vspdData1.Text) & iColSep
				.vspdData1.Col = C_BOMNo1 : txtVal1 = txtVal1 & Trim(.vspdData1.Text) & iColSep
				.vspdData1.Col = C_TrackingNo1 : txtVal1 = txtVal1 & Trim(.vspdData1.Text) & iColSep
				.vspdData1.Col = C_DueDt1 : txtVal1 = txtVal1 & UNIConvDate(.vspdData1.Text) & iColSep
				.vspdData1.Col = C_MpsNo1 : txtVal1 = txtVal1 & Trim(.vspdData1.Text) & iColSep
				.vspdData1.Col = C_SplitSeq1 : txtVal1 = txtVal1 & Trim(.vspdData1.Text) & iRowSep
				TmpArr1(pvCnt) = txtVal1
			end if	
		Next
		
		Redim TmpArr2(0)
		IntQtyRow = 0
		
		For pvCnt = 1 To .vspdData2.MaxRows
				 .vspdData2.Row = pvCnt
		        .vspdData2.CoL = C_Select2
		    If  .vspdData2.Text = "1"   then
				
				txtVal1 = ""
				.vspdData2.Col = C_ItemCd2 : txtVal1 = txtVal1 & Trim(.vspdData2.Text) & iColSep
				.vspdData2.Col = C_BOMNo2 : txtVal1 = txtVal1 & Trim(.vspdData2.Text) & iColSep
				.vspdData2.Col = C_TrackingNo2 : txtVal1 = txtVal1 & Trim(.vspdData2.Text) & iColSep
				.vspdData2.Col = C_DueDt2 : txtVal1 = txtVal1 & UNIConvDate(.vspdData2.Text) & iColSep
				.vspdData2.Col = C_MpsNo2 : txtVal1 = txtVal1 & Trim(.vspdData2.Text) & iColSep
				.vspdData2.Col = C_SplitSeq2 : txtVal1 = txtVal1 & Trim(.vspdData2.Text) & iRowSep
				TmpArr1(pvCnt + .vspdData1.MaxRows) = txtVal1
			
			
			
				.vspdData2.Col = C_SelectForPurQty2 
				If Trim(.vspdData2.Text) = "1" Then
					.vspdData2.Col = C_ItemCd2 : txtVal2 = txtVal2 & Trim(.vspdData2.Text) & iColSep
					.vspdData2.Col = C_BOMNo2 : txtVal2 = txtVal2 & Trim(.vspdData2.Text) & iColSep
					.vspdData2.Col = C_TrackingNo2 : txtVal2 = txtVal2 & Trim(.vspdData2.Text) & iColSep
					.vspdData2.Col = C_DueDt2 : txtVal2 = txtVal2 & UNIConvDate(.vspdData2.Text) & iColSep
					.vspdData2.Col = C_MpsNo2 : txtVal2 = txtVal2 & Trim(.vspdData2.Text) & iColSep
					.vspdData2.Col = C_SplitSeq2 : txtVal2 = txtVal2 & Trim(.vspdData2.Text) & iColSep
					.vspdData2.Col = C_PlanQty2 : txtVal2 = txtVal2 & Trim(.vspdData2.Text) & iRowSep
					
					IntQtyRow = IntQtyRow + 1
					ReDim Preserve TmpArr2(IntQtyRow)
					TmpArr2(IntQtyRow) = txtVal2
				End If
			End IF	
		Next
		
		.txtSpread.value = Join(TmpArr1, "")
		.txtSpread2.value = Join(TmpArr2, "")
		
		Call ExecMyBizASP(frm1, BIZ_PGM_CONFIRM_ID)
		
		'Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    End With
    ConfirmSave = False	
End Function

Function ConfirmOk()
	Dim Index
	
	frm1.vspdData1.ReDraw = false
	frm1.vspdData2.ReDraw = false
	
	ggoSpread.Source = frm1.vspdData1
	ggoSpread.spreadLock C_Select1, -1, C_Select1
	
	frm1.vspdData1.Col = 0
	For Index = 1 To frm1.vspdData1.MaxRows
		frm1.vspdData1.Row = Index
		If frm1.vspdData1.text = ggoSpread.UpdateFlag Then
			ggoSpread.SSDeleteFlag Index
		End If	
	Next	
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.spreadLock C_Select2, -1, C_Select2
	
	frm1.vspdData2.Col = 0
	For Index = 1 To frm1.vspdData2.MaxRows
		frm1.vspdData2.Row = Index
		If frm1.vspdData2.text = ggoSpread.UpdateFlag Then
			ggoSpread.SSDeleteFlag Index
		End If	
	Next	
	
	frm1.vspdData1.ReDraw = True
	frm1.vspdData2.ReDraw = True
	
	frm1.btnAutoSel.disabled = True
	frm1.btnConfirm.disabled = True
	frm1.btnRun.disabled = True
	
End Function

Function JumpToOrder()

    Dim IntRetCd, strVal
	
	WriteCookie "txtPlantCd", UCase(Trim(frm1.txtPlantCd.value))
	WriteCookie "txtPlantNm", frm1.txtPlantNm.value 
	WriteCookie "txtProdOrderNo", UCase(Trim(frm1.txtProdOrderNo.value))
	WriteCookie "txtPGMID", "P4110MA1"
	
	If lstrPgmID <> "" Then
		PgmJump(BIZ_PGM_JUMPTOORDER_ID2)
	Else
		PgmJump(BIZ_PGM_JUMPTOORDER_ID1)
	End If
	
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
    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet("*")                                                    '��: Setup the Spread sheet
  '  Call InitVariables                                                      '��: Initializes local global variables
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    'Call InitComboBox
    Call SetToolBar("1000000000011")										'��: ��ư ���� ���� 
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

'========================================================================================
' Function Name : vspdData1_Click
' Function Desc : �׸��� ��� Ŭ���� ���� 
'========================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
 	
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData1
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
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
    
    Call SetPopupMenuItemInf("0000111111")         'ȭ�麰 ���� 
    
 	Set gActiveSpdSheet = frm1.vspdData2
    
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
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row)
	
	With frm1.vspdData2 

		Select Case Col
			
		    Case C_PlanQty2
				
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.UpdateRow Row
				.Row = Row
				.Col = C_PlanQty2
				If .Value <= 0 Then
					Call DisplayMsgBox("169918", "x", "x", "x")
					.Value = ""
					.Focus
					Set gActiveElement = document.activeElement 
					Exit Sub
				End If
				
				.Col = C_SelectForPurQty2
				.value = 1
							
		End Select
    
   End With

End Sub


Function btnAutoSel_onClick()

	If lgButtonSelection = "SELECT" Then
		lgButtonSelection = "DESELECT"
		frm1.btnAutoSel.value = "��ü����"
	Else
		lgButtonSelection = "SELECT"
		frm1.btnAutoSel.value = "��ü�������"
	End If

	Dim index,Count
	Dim strFlag
	
	frm1.vspdData1.ReDraw = false
	
	Count = frm1.vspdData1.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData1.Row = index
		frm1.vspdData1.Col = C_Select1
		
		strFlag = frm1.vspdData1.Value
		
		If lgButtonSelection = "SELECT" Then
			frm1.vspdData1.Value = 1
			frm1.vspdData1.Col = 0 
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData1.Value = 0
			frm1.vspdData1.Col = 0 
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SSDeleteFlag Index
			frm1.vspdData1.Text=""
		End if

	Next 
	
	frm1.vspdData1.ReDraw = true
	
	frm1.vspdData2.ReDraw = false
	
	Count = frm1.vspdData2.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData2.Row = index
		frm1.vspdData2.Col = C_Select2
		
		strFlag = frm1.vspdData2.Value
		
		If lgButtonSelection = "SELECT" Then
			frm1.vspdData2.Value = 1
			frm1.vspdData2.Col = 0 
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.UpdateRow Index
		Else
			frm1.vspdData2.Value = 0
			frm1.vspdData2.Col = 0 
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SSDeleteFlag Index
			frm1.vspdData2.Text=""
		End if

	Next 
	
	frm1.vspdData2.ReDraw = true

End Function

'========================================================================================
' Function Name : vspdData1_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData1.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub
 
'========================================================================================
' Function Name : vspdData2_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData2.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'==========================================================================================
'   Event Name : vspdData_DragDropBlock
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_DragDropBlock(ByVal Col , ByVal Row , ByVal Col2 , ByVal Row2 , ByVal NewCol , ByVal NewRow , ByVal NewCol2 , ByVal NewRow2 , ByVal Overwrite , Action , DataOnly , Cancel )
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	With frm1.vspdData1
		.Row = Row
		.Col = C_Select1
		
		ggoSpread.Source = frm1.vspdData1
		
		If ButtonDown = 1 Then
			ggoSpread.UpdateRow Row
		Else
			ggoSpread.SSDeleteFlag Row,Row
		End If
		
	End With
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
	With frm1.vspdData2
		.Row = Row
		.Col = C_Select2

		ggoSpread.Source = frm1.vspdData2
		
		If ButtonDown = 1 Then
			.Col = C_ProcurType2
			If Trim(.Text) ="P" Then
				ggoSpread.SpreadUnLock C_PlanQty2, Row, C_PlanQty2,Row
				ggoSpread.SSSetRequired C_PlanQty2, Row, Row
			End If	
			ggoSpread.UpdateRow Row	
		Else
			.Col = C_ProcurType2
			If Trim(.Text) ="P" Then
				ggoSpread.SpreadLock C_PlanQty2, Row, C_PlanQty2,Row
				ggoSpread.SSSetProtected C_PlanQty2, Row, Row
			End If	
			.Col = C_SelectForPurQty2
			If Trim(.Text) = "1" And frm1.vspdData2.MaxRows > 1 Then
				ggoSpread.EditUndo
			End If
			ggoSpread.SSDeleteFlag Row,Row
		End If
		
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
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			 If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'��: �ٸ� ����Ͻ� ������ ���� ���̸� �� �̻� ��������ASP�� ȣ������ ���� 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey2 <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call LayerShowHide(1)
			If DbQuery = False Then	
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
' Function Name : vspdData2_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
End Sub 

'========================================================================================
' Function Name : vspdData1_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData2_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData1_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

	'If NewCol = C_XXX or Col = C_XXX Then
	'	Cancel = True
	'	Exit Sub
	'End If

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : vspdData1_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
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
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtExecFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtExecFromDt.Action = 7
        SetFocusToDocument("M")
		Frm1.txtExecFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt_OnBlur()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtExecFromDt_OnBlur()
	Dim DtInvCloseDt
	Dim DtExecFromDt
	Dim DtStartDt
	If frm1.txtExecFromDt.text = "" Then Exit Sub
	
	DtInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)
	DtExecFromDt = UniConvDateAToB(frm1.txtExecFromDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	DtStartDt    = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If DtExecFromDt <= DtInvCloseDt Then
		lgDateCheckFlg = "False1"
		Call DisplayMsgBox("189250", "x", "x", "x")
		frm1.txtExecFromDt.text = UNIDateAdd ("D", 1, lgInvCloseDt, parent.gDateFormat)
		frm1.txtExecFromDt.focus
		Set gActiveElement = document.activeElement
		Exit Sub
'	ElseIf DtExecFromDt > DtStartDt Then
'		If DtStartDt > DtInvCloseDt Then
'			lgDateCheckFlg = "False2"
'			Call DisplayMsgBox("189251", "x", "x", "x")
''			frm1.txtExecFromDt.text = ""
'			frm1.txtExecFromDt.focus
'			Set gActiveElement = document.activeElement
'			Exit Sub
'		Else
'			lgDateCheckFlg = ""
'		End If
	Else
		lgDateCheckFlg = ""
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

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If	
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables
    															'��: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function																'��: Query db data

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
    Call parent.FncExport(parent.C_MULTI)                                                   '��: Protect system from crashing
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
	FncExit = True
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
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '��: Protect system from crashing
    
    With frm1
  
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'��: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtPlanOrderNo=" & Trim(.txtPlanOrderNo.value)				'��: ��ȸ ���� ����Ÿ		
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
  
	
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
	Call SetToolBar("1000000000111")										'��: ��ư ���� ���� 
	If lgIntFlgMode = parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True	
	frm1.btnAutoSel.disabled = False
End Function



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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenErrorList()">ERROR��������Ʈ</A>	<A href="vbscript:OpenRef()">Ȯ�������ȸ</A></TD>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14" ALT="����">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>			 						
									<TD CLASS=TD5 NOWRAP>����������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="14" ALT="����������ȣ"></TD>
									<TD CLASS=TD5 NOWRAP>�������ݿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME=chkInvStock ALT="�������ݿ�" tag="11" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid"></INPUT>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="14" ALT="ǰ��">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>��ȹ������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlanOrderNo" SIZE=18 tag="14" ALT="��ȹ������ȣ"></TD>
									<TD CLASS=TD5 NOWRAP>�������ݿ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME=chkSFStock ALT="�������ݿ�" tag="11" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid"></INPUT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�԰�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpecification" SIZE=40 MAXLENGTH=50 tag="14" ALT="�԰�"></TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtOrderQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="14X3" ALT="��������" MAXLENGTH="15" SIZE="10" id=OBJECT1></OBJECT>');</SCRIPT>
										</TD>
									
									<TD CLASS=TD5 NOWRAP>Forward</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME=chkForward ALT="Forward" tag="11" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid"></INPUT></TD>
								</TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>�۾�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtStartDt CLASSID=<%=gCLSIDFPDT%> ALT="������" tag="14" ></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEndDt CLASSID=<%=gCLSIDFPDT%> ALT="������" tag="14" > </OBJECT>');</SCRIPT>									
									</TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrderUnit" SIZE=5 MAXLENGTH=3 tag="14xxxU" ALT="������"></TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime3 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtExecFromDt CLASSID=<%=gCLSIDFPDT%> ALT="��������" tag="12X1"> </OBJECT>');</SCRIPT></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=50% valign=top>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData1 ID = "A" width="100%" tag="2" TITLE="SPREAD" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
				<TR>	
					<TD WIDTH=100% HEIGHT=50% valign=top>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 ID = "B" width="100%" tag="2" TITLE="SPREAD" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					   <TD WIDTH = 10>&nbsp;</TD>		              
		               <TD><BUTTON NAME="btnRun" ONCLICK="vbscript:Run()" CLASS="CLSMBTN">����</BUTTON>&nbsp;<a><button name="btnAutoSel" class="clsmbtn">��ü����</button></a>&nbsp;<BUTTON NAME="btnConfirm" ONCLICK="vbscript:Confirm()" CLASS="CLSMBTN">��ȯ</BUTTON><TD WIDTH=* Align=right><A href="vbscript:JumpTOOrder">������������</A></TD></TD>
	                </TR>
	            </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hPlanOrderNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hStartDt" tag="24"><INPUT TYPE=HIDDEN NAME="hEndDt" tag="24"><INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
