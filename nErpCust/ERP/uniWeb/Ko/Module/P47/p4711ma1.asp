<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4711ma1
'*  4. Program Name         : Resource Consumption (Batch)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park, BumSoo
'* 10. Modifier (Last)      : Park, BumSoo
'* 11. Comment              :
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->							<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ��� -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'======================================================================================================select * from b_message====-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ��� -->

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_SHIFT		= "p4711mb1.asp"										 '��: �����Ͻ� ���� ASP�� 
Const BIZ_EXECUTE_ID	= "p4711mb2.asp"
Const BIZ_CANCEL_ID		= "p4711mb3.asp"
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim strDate 
Dim StartDate 
Dim strYear
Dim strMonth
Dim strDay
Dim BaseDate

BaseDate = "<%=GetsvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

Dim IsOpenPop
Dim lgShiftCnt
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
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
End Sub

'=============================== 2.1.2 \fTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029() 
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
End Sub 

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
'==========================================  2.2.1 SetDefaultVal()  ==========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtReportDtFrom.text = StartDate
	frm1.txtReportDtTo.text   = StrDate
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Dim i

	For i = lgShiftCnt To 1 Step -1
		frm1.cboShiftCdFrom.remove(i)
		frm1.cboShiftCdTo.remove(i)  
	Next

    Dim strVal
	
	strVal = BIZ_PGM_SHIFT & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	
    Call RunMyBizASP(MyBizASP, strVal)
	
End Sub

'==========================================  2.2.6 InitStatusCombo()  =======================================
'	Name : InitStatusCombo()
'	Description : Combo Display
'========================================================================================================= 
Sub InitStatusCombo()
	Call SetCombo(frm1.cboStatus, "R", "�����")
	Call SetCombo(frm1.cboStatus, "C", "��ҵ�")		'��: InitCombo ���� �ؾ� �Ǵµ� �ӽ÷� ���� ���� 
End Sub


'******************************************  2.3 Operation ó���Լ�  *************************************
'	���: Operation ó���κ� 
'	����: Tabó��, Reference���� ���Ѵ�. 
'********************************************************************************************************* 

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

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
   	arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
    
    arrHeader(0) = "����"							' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBatchRunNo()  -------------------------------------------------
'	Name : OpenBatchRunNo()
'	Description : Batch Run No. PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBatchRunNo()

	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtBatchRunNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = UCase(Trim(frm1.txtPlantCd.value))
	arrParam(1) = Trim(frm1.txtPlantNm.value)
	arrParam(2) = UCase(Trim(frm1.txtBatchRunNo.value))

	iCalledAspName = AskPRAspName("p4711pa1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4711pa1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBatchRunNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtBatchRunNo.focus
	
End Function

'------------------------------------------  OpenProdOrderNoFrom()  -------------------------------------------------
'	Name : OpenProdOrderNoFrom()
'	Description : Condition Production Order PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenProdOrderNoFrom()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	If IsOpenPop = True Or UCase(frm1.txtProdtOrderNoFrom.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrderNoFrom.value)
	arrParam(6) = ""
	arrParam(7) = Trim(frm1.txtItemCdFrom.value)
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
		Call SetProdOrderNoFrom(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrderNoFrom.focus
		
End Function

'------------------------------------------  OpenProdOrderNoTo()  -------------------------------------------------
'	Name : OpenProdOrderNoTo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNoTo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	If IsOpenPop = True Or UCase(frm1.txtProdtOrderNoTo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = ""
	arrParam(2) = ""
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrderNoTo.value)
	arrParam(6) = ""
	arrParam(7) = Trim(frm1.txtItemCdTo.value)
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
		Call SetProdOrderNoTo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrderNoTo.focus
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCdFrom()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtItemCdFrom.Value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
    iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCdFrom(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCdFrom.focus

End Function
'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemCdTo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCdTo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtItemCdTo.value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"   
	
	iCalledAspName = AskPRAspName("b1b11pa3")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
	End If
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCdTo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCdTo.focus

End Function

'------------------------------------------  OpenWcCdFrom()  ------------------------------------------------
'	Name : OpenWcCdFrom()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCdFrom()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCdFrom.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNmFrom.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")		' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCdFrom(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCdFrom.focus
	
End Function

'------------------------------------------  OpenWcCdTo()  ------------------------------------------------
'	Name : OpenWcCdTo()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCdTo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtWCCdTo.Value)								' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNmTo.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCdTo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCdTo.focus
	
End Function

'------------------------------------------  OpenErrorRef()  -------------------------------------------------
'	Name : OpenErrorRef()
'	Description : Part Reference PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenErrorRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If

	If frm1.txtBatchRunNo.value= "" Then
		Call DisplayMsgBox("971012","X", "�̷¹�ȣ","X")
		frm1.txtBatchRunNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	arrParam(0) = Trim(UCase(frm1.txtPlantCd.value))	'��: ��ȸ ���� ����Ÿ 
	arrParam(1) = Trim(frm1.txtPlantNm.value)			'��: ��ȸ ���� ����Ÿ 
	arrParam(2) = Trim(frm1.txtBatchRunNo.value)		'��: ��ȸ ���� ����Ÿ 
	
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("p4711ra2")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4711ra2", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

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
	Call InitComboBox
End Function

'------------------------------------------  SetBatchRunNo()  --------------------------------------------------
'	Name : SetBatchRunNo()
'	Description : ResourceGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBatchRunNo(byval arrRet)
	frm1.txtBatchRunNo.Value = arrRet(0)
	frm1.cboStatus.Value	 = arrRet(1)
	frm1.txtSuccessCnt.Value = arrRet(2)
	frm1.txtErrorCnt.Value	 = arrRet(3)
End Function

'------------------------------------------  SetFrProdOrderNo()  -------------------------------------------------
'	Name : SetFrProdOrderNo()
'	Description : ProdOrderNo Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNoFrom(ByVal arrRet)
	frm1.txtProdtOrderNoFrom.value = arrRet(0) 
End Function

'------------------------------------------  SetProdOrderNoTo()  -------------------------------------------------
'	Name : SetProdOrderNoTo()
'	Description : ProdOrderNo Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNoTo(ByVal arrRet)
	frm1.txtProdtOrderNoTo.value = arrRet(0) 
End Function

'------------------------------------------  SetItemCdFrom()  -------------------------------------------------
'	Name : SetItemCdFrom()
'	Description : ProdOrderNo Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCdFrom(ByVal arrRet)
	frm1.txtItemCdFrom.value = arrRet(0)
	frm1.txtItemNmFrom.value = arrRet(1)  
End Function

'------------------------------------------  SetItemCdTo()  -------------------------------------------------
'	Name : SetItemCdTo()
'	Description : ProdOrderNo Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetItemCdTo(ByVal arrRet)
	frm1.txtItemCdTo.value = arrRet(0)
	frm1.txtItemNmTo.value = arrRet(1)  
End Function

'------------------------------------------  SetWcCdFrom()  -------------------------------------------------
'	Name : SetWcCdFrom()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCdFrom(byval arrRet)
	frm1.txtWCCdFrom.Value    = arrRet(0)		
	frm1.txtWCNmFrom.Value    = arrRet(1)		
End Function

'------------------------------------------  SetWcCdTo()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCdTo(byval arrRet)
	frm1.txtWCCdTo.Value    = arrRet(0)		
	frm1.txtWCNmTo.Value    = arrRet(1)		
End Function

Sub txtPlantCd_OnChange()
    If frm1.txtPlantCd.value <> "" Then
		Call InitComboBox	
	End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'++++++++++++++++++++++++++++++++++++++++++  2.5.2 Execute  +++++++++++++++++++++++++++++++++++++++
'        Name : Execute()
'        Description : MRP ���� Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function Execute()

	Dim strVal
		
    Err.Clear															'��: Protect system from crashing
    Execute = False														'��: Processing is NG

    If Not chkField(Document, "1") Then									'��: Check contents area
       Exit Function
    End If
	
	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If ValidDateCheck(frm1.txtReportDtFrom, frm1.txtReportDtTo) = False Then Exit Function
        
    Call LayerShowHide(1)
    
	strVal = BIZ_EXECUTE_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtProdtOrderNoFrom=" & Trim(frm1.txtProdtOrderNoFrom.value)
	strVal = strVal & "&txtProdtOrderNoTo=" & Trim(frm1.txtProdtOrderNoTo.value)
	strVal = strVal & "&txtItemCdFrom=" & Trim(frm1.txtItemCdFrom.value)
	strVal = strVal & "&txtItemCdTo=" & Trim(frm1.txtItemCdTo.value)
	strVal = strVal & "&txtWcCdFrom=" & Trim(frm1.txtWcCdFrom.value)
	strVal = strVal & "&txtWcCdTo=" & Trim(frm1.txtWcCdTo.value)
	strVal = strVal & "&cboShiftCdFrom=" & Trim(frm1.cboShiftCdFrom.value)
	strVal = strVal & "&cboShiftCdTo=" & Trim(frm1.cboShiftCdTo.value)
	strVal = strVal & "&txtReportDtFrom=" & Trim(frm1.txtReportDtFrom.text)
	strVal = strVal & "&txtReportDtTo=" & Trim(frm1.txtReportDtTo.text)
	
	Call RunMyBizASP(MyBizASP, strVal)
	
    Execute = True 
            
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5.2 Cancel()  +++++++++++++++++++++++++++++++++++++++
'        Name : Cancel()
'        Description : MRP ���� Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function Cancel()

	Dim strVal
		
    Err.Clear															'��: Protect system from crashing
    Cancel = False														'��: Processing is NG

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtBatchRunNo.value= "" Then
		Call DisplayMsgBox("971012","X", "�̷¹�ȣ","X")
		frm1.txtBatchRunNo.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If

	Dim IntRetCD
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    Call LayerShowHide(1)
    
	strVal = BIZ_CANCEL_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtBatchRunNo=" & Trim(frm1.txtBatchRunNo.value)
    Call RunMyBizASP(MyBizASP, strVal)
	
    Cancel = True 
            
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
    Call SetDefaultVal
    Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call SetDefaultVal																	'��: Initializes local global variables
    Call InitVariables
	Call InitStatusCombo
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm	
		Call InitComboBox()
		frm1.txtProdtOrderNoFrom.focus 
	ELSE
		frm1.txtPlantCd.focus 
	End If   
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtReportDtFrom_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDtFrom_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportDtFrom.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtReportDtFrom.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtReportDtTo_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReportDtTo_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportDtTo.Action = 7 
		Call SetFocusToDocument("M")
		Frm1.txtReportDtTo.Focus
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
	On Error Resume Next       
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
     On Error Resume Next                                                   '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()                                                   '��: Protect system from crashing
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
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                    '��:ȭ�� ����, Tab ���� 
End Function

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
    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================

Function DbSave() 
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()															'��: ���� ������ ���� ���� 
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ��Һ���(Batch)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�̷¹�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBatchRunNo" SIZE=18 MAXLENGTH=18 tag="11XXXU"  ALT="�̷¹�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBatchRunNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBatchRunNo()" >&nbsp;<SELECT NAME="cboStatus" ALT="Status" STYLE="Width: 98px;" tag="14"></SELECT></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>���ȵȽ�����</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtSuccessCnt" SIZE=16 MAXLENGTH=16 tag="14xxxU" ALT="���ȵȽ�����"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>ERROR��</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtErrorCnt" SIZE=16 MAXLENGTH=16 tag="14xxxU" ALT="ERROR��"></TD>
							</TR>
							<TR>
							    <TD CLASS="TD5" NOWRAP>����������ȣ</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtProdtOrderNoFrom" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNoFrom()">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtProdtOrderNoTo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="����������ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenProdOrderNoTo()">
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCdFrom" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCdFrom" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCdFrom()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmFrom" SIZE=40 MAXLENGTH=40 tag="14" ALT="ǰ���">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCdTo" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCdTo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCdTo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmTo" SIZE=40 MAXLENGTH=40 tag="14" ALT="ǰ���">&nbsp;
								</TD>
							</TR>
							<TR>								
								<TD CLASS=TD5 NOWRAP>�۾���</TD>
								<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtWCCdFrom" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCdFrom()"> <INPUT TYPE=TEXT  NAME="txtWCNmFrom" SIZE=40 MAXLENGTH=40 tag="14">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtWCCdTo" SIZE=7 MAXLENGTH=7 tag="11xxxU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenWcCdTo()"> <INPUT TYPE=TEXT  NAME="txtWCNmTo" SIZE=40 MAXLENGTH=40 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Shift</TD>
								<TD CLASS=TD6 NOWRAP>
								<SELECT NAME="cboShiftCdFrom" ALT="���� Shift" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
								&nbsp;~&nbsp;
								<SELECT NAME="cboShiftCdTo" ALT="���� Shift" STYLE="Width: 98px;" tag="11"><OPTION VALUE=""></OPTION></SELECT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���������</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4711ma1_I498841530_txtReportDtFrom.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p4711ma1_I689011383_txtReportDtTo.js'></script>
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
					<TD><BUTTON NAME="btnExecute" CLASS="CLSMBTN" Flag=1 onclick="Execute()">����</BUTTON>&nbsp;<BUTTON NAME="btnCancel" CLASS="CLSMBTN" Flag=1 onclick="Cancel()">���</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
