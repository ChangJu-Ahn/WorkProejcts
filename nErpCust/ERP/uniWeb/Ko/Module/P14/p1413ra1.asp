
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name          : Production																*
'*  2. Function Name        :																			*
'*  3. Program ID           : p1413ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Mass Replacement															*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2002/03/14																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : RYU SUNG WON																*
'* 10. Modifier (Last)      : 																			*
'* 11. Comment              :																			*
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--####################################################################################################
'#						1. �� �� ��																		#
'#####################################################################################################-->

<!--********************************************  1.1 Inc ����  *****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--============================================  1.1.1 Style Sheet  ====================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--============================================  1.1.2 ���� Include  ===================================
'=====================================================================================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

'********************************************  1.2 Global ����/��� ����  *******************************
'*	Description : 1. Constant�� �ݵ�� �빮�� ǥ��														*
'********************************************************************************************************
Const BIZ_PGM_QRY_ID = "p1413rb1.asp"						'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_LOOKUP_REASON_INFO = "p1412rb2.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_LOOKUP_ECN_INFO	= "p1412rb3.asp"
'============================================  1.2.1 Global ��� ����  ==================================
'========================================================================================================
	
'==========================================  1.2.0 Common variables =====================================
'	1. Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
	
'============================================  1.2.2 Global ���� ����  ==================================
'========================================================================================================
Dim arrReturn
Dim lgPlantCD
Dim strFromStatus
Dim strToStatus
Dim strThirdStatus
Dim IsOpenPop
Dim arrParent
Dim lgBlnEcnValueChanged		'txtEcnNo�� onChange �̺�Ʈ�� �Ϸ���� Click�̺�Ʈ�� ����Ǳ� ���� Flag
Dim lgBlnReasonValueChanged		'txtReasonCd�� onChange �̺�Ʈ�� �Ϸ���� Click�̺�Ʈ�� ����Ǳ� ���� Flag
Dim lgBlnReasonChangeEventFinished

ArrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName
'============================================  1.2.3 Global Variable�� ����  ============================
'========================================================================================================
'----------------  ���� Global ������ ����  -------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						2. Function ��																	#
'#																										#
'#	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� ���					#
'#	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.							#
'#						 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����)			#
'########################################################################################################
'*******************************************  2.1 ���� �ʱ�ȭ �Լ�  *************************************
'*	���: �����ʱ�ȭ																					*
'*	Description : Global���� ó��, �����ʱ�ȭ ���� �۾��� �Ѵ�.											*
'********************************************************************************************************

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)				=
'========================================================================================================
Function InitVariables()
	lgIntFlgMode = PopupParent.OPMD_CMODE								'Indicates that current mode is Create mode	
	Self.Returnvalue = Array("")
End Function

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================%>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "P", "NOCOOKIE","RA") %>
	<% Call loadBNumericFormatA("I", "P", "NOCOOKIE","RA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim LocSvrDate
	LocSvrDate = "<%=GetSvrDate%>"
	frm1.txtValidFromDt.text= UniConvDateAToB(LocSvrDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
	frm1.txtValidToDt.text	= UniConvDateAToB("2999-12-31", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
End Sub

'==========================================   2.1.2 InitSetting()   =====================================
'=	Name : InitSetting()																				=
'=	Description : Passed Parameter�� Variable�� Setting�Ѵ�.											=
'========================================================================================================
Function InitSetting()
	Dim ArgArray						<%'Arguments�� �Ѱܹ��� Array%>
	
	ArgArray  = ArrParent(1)

	frm1.txtPlantCd.value = ArgArray(0)	'plant cd
	frm1.txtItemCd.value = ArgArray(1)	'item cd
	frm1.txtBomType.value = ArgArray(2)	'bom type
	
	frm1.cboSupplyType.selectedIndex = 1	'���� 

	lgBlnEcnValueChanged = True
	lgBlnReasonValueChanged = True
	lgBlnReasonChangeEventFinished = True
End Function

'==========================================   2.1.3 InitComboBox()  =====================================
'=	Name : InitComboBox()																				=
'=	Description : ComboBox�� Value�� Setting�Ѵ�.														=
'========================================================================================================
Sub InitComboBox()
	Dim strCbo
    Dim strCboCd
    Dim strCboNm

	'****************************
    'List Minor code(�����󱸺�)
    '****************************
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2201", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	strCboCd = Replace(lgF0,chr(11),vbTab)
    strCboNm = Replace(lgF1,chr(11),vbTab)
    
    Call SetCombo2(frm1.cboSupplyType, lgF0, lgF1, Chr(11))
End Sub

'*******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  *************************************
'*	���: ȭ���ʱ�ȭ																					*
'*	Description : ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�.						*
'********************************************************************************************************

'++++++++++++++++++++++++++++++++++++++++++  2.3 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	������ ���� Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"					' �˾� ��Ī 
	arrParam(1) = "B_PLANT"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"						' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"						' Field��(0)
    arrField(1) = "PLANT_NM"						' Field��(1)
    
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenBomNo()  -------------------------------------------------
'	Name : OpenBomNo()
'	Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBomNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then Exit Function

	If UCase(Trim(frm1.txtPlantCd.value)) = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "ǰ��", "X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "BOM�˾�"						' �˾� ��Ī 
	arrParam(1) = "B_MINOR"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtBomType.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1401", "''", "S") & " "
	arrParam(5) = "BOM Type"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "BOM Type"					' Header��(0)
    arrHeader(1) = "BOM Ư��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetBomNo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	frm1.txtBomType.focus
	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(6), arrField(10)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "����", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)								' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)							' Item Code
	arrParam(2) = ""												' Combo Set Data:"1029!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""													' Default Value
		
	arrField(0) = 1		'ITEM_CD
    arrField(1) = 2 	'ITEM_NM											
    arrField(2) = 5		'ITEM_ACCT
    arrField(3) = 9 	'PROC_TYPE
    arrField(4) = 4 	'BASIC_UNIT
    arrField(5) = 51	'SINGLE_ROUT_FLG
    arrField(6) = 52	'Major_Work_Center
    arrField(7) = 13	'Phantom_flg
    arrField(8) = 18	'valid_from_dt
    arrField(9) = 19	'valid_to_dt
    arrField(10) = 3	' Field��(1) : "SPECIFICATION"   
    
	iCalledAspName = AskPRAspName("B1B11PA4")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	frm1.txtItemCd.focus
		
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit(ByVal pTarget)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iUnit

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	If pTarget = "CHILD" Then
		iUnit = frm1.txtChildUnit.value
	ElseIF pTarget = "PRNT" Then
		iUnit = frm1.txtPrntUnit.value
	End If
	
	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(iUnit)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "����"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "������"
   
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet, pTarget)
	End If	
	
	Call SetFocusToDocument("P")
	If pTarget = "CHILD" Then
		frm1.txtChildUnit.focus
	ElseIF pTarget = "PRNT" Then
		frm1.txtPrntUnit.focus
	End If 
	
End Function

'------------------------------------------  OpenECNInfo()  ----------------------------------------------
'	Name : OpenECNInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenECNInfo()

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtECNNo.className) = UCase(PopupParent.UCN_PROTECTED) Then
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtECNNo.value)	' ECNNo
	arrParam(1) = ""						' ReasonCd
	arrParam(2) = ""						' Status
	arrParam(3) = ""						' EBomFlg
	arrParam(4) = ""						' MBomFlg

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetECNInfo(arrRet)
	End If
	
	Call SetFocusToDocument("P")
	frm1.txtECNNo.focus
	
End Function

'------------------------------------------  OpenReasonPopup()  ------------------------------------------
'	Name : OpenReasonPopup()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonPopup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtReasonCd.className) = UCase(PopupParent.UCN_PROTECTED) Then
		Exit Function
	End If
    
	'---------------------------------------------
	' Parameter Setting
	'--------------------------------------------- 

	IsOpenPop = True

	arrParam(0) = "ECN ��ȣ�˾�"					' �˾� ��Ī 
	arrParam(1) = "B_MINOR"								' TABLE ��Ī 
	arrParam(2) = UCase(Trim(frm1.txtReasonCd.value))	' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = "MAJOR_CD = " & FilterVar("P1402", "''", "S") & ""
	
	arrParam(5) = "����ٰ�"						' TextBox ��Ī 
	
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
        
    arrHeader(0) = "����ٰ�"					' Header��(0)
    arrHeader(1) = "����ٰŸ�"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReasonInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("P")
	frm1.txtECNNo.focus
	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	
	frm1.txtItemCd.value = UCase(Trim(arrRet(0)))
	frm1.txtItemNm.value =	Trim(arrRet(1))
		
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlantCd(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)
	frm1.txtPlantNm.Value    = arrRet(1)				
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
'	Name : SetBomNo()
'	Description : Bom No Popup���� return�� �� 
'--------------------------------------------------------------------------------------------------------- 
Function SetBomNo(byval arrRet)
		frm1.txtBomType.Value	= arrRet(0) 		
End Function

'------------------------------------------  SetUnit()  ------------------------------------------------
'	Name : SetUnit()
'	Description : Open Unit Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetUnit(ByVal arrRet, ByVal pTarget)
	If pTarget = "CHILD" Then
		frm1.txtChildUnit.value = arrRet(0)
	ElseIF pTarget = "PRNT" Then
		frm1.txtPrntUnit.value = arrRet(0)
	End If
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetECNInfo()  ------------------------------------------------
'	Name : SetECNInfo()
'	Description : ECNInfo Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetECNInfo(byval arrRet)
	frm1.txtEcnNo.Value    = arrRet(0)		
	frm1.txtEcnDesc.Value  = arrRet(1)
	frm1.txtReasonCd.Value = arrRet(2)
	frm1.txtReasonNm.value = arrRet(3)
	
	lgBlnReasonValueChanged = True
End Function

'------------------------------------------  SetReasonInfo()  --------------------------------------------------
'	Name : SetReasonInfo()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function SetReasonInfo(byval arrRet)
	frm1.txtReasonCd.Value			= arrRet(0)	
	frm1.txtReasonNm.Value			= arrRet(1)
	
	lgBlnFlgChgValue = True
	lgBlnReasonValueChanged = True
End Function
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If Not chkField(Document, "2") Then									
       Exit Function
    End If

	If lgBlnEcnValueChanged = False Then Exit Function
	If lgBlnReasonValueChanged = False Then 
		If lgBlnReasonChangeEventFinished = True Then
			Call DisplayMsgBox("182803", vbOKOnly, "", "")
		End If
		frm1.txtReasonCd.focus
		Exit Function
	End If
    
    If UniConvNum(frm1.txtChildItemQty.Text, 0) = 0 Then
		Call DisplayMsgBox("970022", VBOKOnly, "��ǰ����ؼ�", "0")
		frm1.txtChildItemQty.focus
		Exit Function
	End If

    If UniConvNum(frm1.txtPrntItemQty.Text, 0) = 0 Then
		Call DisplayMsgBox("970022", VBOKOnly, "��ǰ����ؼ�", "0")
		frm1.txtPrntItemQty.focus
		Exit Function
	End If
	
	If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function

	Redim arrReturn(20)
		
	arrReturn(0) = UCase(Trim(frm1.txtPlantCd.value))
	arrReturn(1) = UCase(Trim(frm1.hItemCd.value))
	arrReturn(2) = UCase(Trim(frm1.txtItemNm.value))
	arrReturn(3) = UCase(Trim(frm1.txtBomType.value))
	arrReturn(4) = UCase(Trim(frm1.txtAcct.value))
	arrReturn(5) = UCase(Trim(frm1.txtSpec.value))
	arrReturn(6) = UCase(Trim(frm1.txtProcurType.value))
	arrReturn(7) = UCase(Trim(frm1.cboSupplyType.value))
	arrReturn(8) = frm1.txtChildItemQty.Text
	arrReturn(9) = UCase(Trim(frm1.txtChildUnit.value))
	arrReturn(10) = frm1.txtPrntItemQty.Text
	arrReturn(11) = UCase(Trim(frm1.txtPrntUnit.value))
	arrReturn(12) = frm1.txtSafetyLt.Text
	arrReturn(13) = frm1.txtLossRate.Text
	arrReturn(14) = frm1.txtValidFromDt.Text
	arrReturn(15) = frm1.txtValidToDt.Text
	arrReturn(16) = UCase(Trim(frm1.txtEcnNo.value))
	arrReturn(17) = frm1.txtEcnDesc.value
	arrReturn(18) = UCase(Trim(frm1.txtReasonCd.value))
	arrReturn(19) = frm1.txtReasonNm.value
	arrReturn(20) = frm1.txtRemark.value

	Self.Returnvalue = arrReturn
	Self.Close()

End Function
	
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	self.close()
End Function
'=========================================  2.3.3 Mouse Pointer ó�� �Լ� ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'==========================================================================================
'   Event Name : LookUpEcnInfo
'   Event Desc : EcnNo Change Event�߻��� ��ȸ 
'==========================================================================================
Sub LookUpEcnInfo()
	Dim strVal
	Dim strEcnNo

	If   LayerShowHide(1) = False Then Exit Sub
	
	lgBlnEcnValueChanged = False
	
	strEcnNo = Trim(frm1.txtEcnNo.value)

	strVal = BIZ_PGM_LOOKUP_ECN_INFO & "?txtMode=" & PopupParent.UID_M0001
	strVal = strVal & "&txtEcnNo=" & Trim(strEcnNo)

	Call RunMyBizASP(MyBizASP, strVal)
End Sub

Sub LookUpEcnInfoOk(ByVal pResult)
	If CBool(pResult) = True Then
		Call ggoOper.SetReqAttr(frm1.txtECNDesc, "Q")
		Call ggoOper.SetReqAttr(frm1.txtReasonCd, "Q")
		lgBlnReasonValueChanged = True
	Else	'Data Not Found
		Call ggoOper.SetReqAttr(frm1.txtECNDesc, "N")
		Call ggoOper.SetReqAttr(frm1.txtReasonCd, "N")
		lgBlnReasonValueChanged = False
	End If

	lgBlnEcnValueChanged = True
End Sub

'==========================================================================================
'   Event Name : LookUpReasonInfo
'   Event Desc : 
'==========================================================================================
Function LookUpReasonInfo()
	Dim strVal
	Dim strReasonCd

	strReasonCd = Trim(frm1.txtReasonCd.value)
	
	If LayerShowHide(1) = False Then Exit Function
	
	lgBlnReasonValueChanged = False
	
	strVal = BIZ_PGM_LOOKUP_REASON_INFO & "?txtMode=" & PopupParent.UID_M0001		<%'��: �����Ͻ� ó�� ASP�� ���� %>
	strVal = strVal & "&txtReasonCd=" & Trim(strReasonCd)

	Call RunMyBizASP(MyBizASP, strVal)

End Function

Sub LookUpReasonInfoOk()
	lgBlnReasonValueChanged = True
	lgBlnReasonChangeEventFinished = True
End Sub

'*******************************************  2.4 POP-UP ó���Լ�  **************************************
'*	���: POP-UP																						*
'*	Description : POP-UP Call�ϴ� �Լ� �� Return Value setting ó��										*
'********************************************************************************************************
'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtValidToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtValidToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtValidToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtECNNo_OnChange
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtECNNo_OnChange()
	Call LookUpEcnInfo()
End Sub

'=======================================================================================================
'   Event Name : txtReasonCd_OnChange
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtReasonCd_OnChange()
	lgBlnReasonChangeEventFinished = False
	Call LookUpReasonInfo()
End Sub

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================



'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================



'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  ++++++++++++++++++++++++++++++++++++++
'+	���� ���α׷����� �ʿ��� ������ ���� Procedure(Sub, Function, Validation & Calulation ���� �Լ�)	+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'########################################################################################################
'#						3. Event ��																		#
'#	���: Event �Լ��� ���� ó��																		#
'#	����: Windowó��, Singleó��, Gridó�� �۾�.														#
'#		  ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.								#
'#		  �� Object������ Grouping�Ѵ�.																	#
'########################################################################################################
'********************************************  3.1 Windowó��  ******************************************
'*	Window�� �߻� �ϴ� ��� Even ó��																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ�				=
'========================================================================================================
Sub Form_Load()
	Call AppendNumberPlace("6", "3", "0")
	Call AppendNumberPlace("7", "2", "2")
	Call AppendNumberPlace("8", "11", "6")
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
	
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029											'��: Load table , B_numeric_format		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    		
	Call ggoOper.LockField(Document, "N")						<% '��: Lock  Suitable  Field %>
	Call SetDefaultVal()
	Call InitVariables()
	Call InitComboBox()
	Call InitSetting()
	
	If frm1.txtPlantCd.value <> "" Then
		frm1.txtItemCd.focus
	End If

End Sub
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'*********************************************  3.2 Tag ó��  *******************************************
'*	Document�� TAG���� �߻� �ϴ� Event ó��																*
'*	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ�							*
'*	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
	Dim IntRetCD 
    FncQuery = False                                                        

    Call ggoOper.ClearField(Document, "2")
    Call SetDefaultVal()
    Call InitVariables
    'Call InitSetting()
	'-----------------------
    'Check condition area
    '----------------------- 

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then
		Exit Function
	End If
       
    FncQuery = True	
End Function

'########################################################################################################
'#					     4. Common Function��															#
'########################################################################################################
'########################################################################################################
'#						5. Interface ��																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
    DbQuery = False                                                         
    
    LayerShowHide(1)							
    
    Dim strVal
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001						
    strVal = strVal & "&txtPlantCd="	& UCase(Trim(frm1.txtPlantCd.value))
    strVal = strVal & "&txtItemCd="		& UCase(Trim(frm1.txtItemCd.value))
    strVal = strVal & "&txtBomType="	& UCase(Trim(frm1.txtBomType.value))
	strVal = strVal & "&PrevNextFlg="	& ""

	Call RunMyBizASP(MyBizASP, strVal)										

    DbQuery = True    
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()															'��: ��ȸ ������ ������� 
    Dim LayerN1
	frm1.hPlantCd.value = frm1.txtPlantCd.value		'CHECK - MB1���� �Ұ����� ��� 
    
	Set LayerN1 = window.document.all("MousePT").style
	
    lgIntFlgMode = PopupParent.OPMD_UMODE											
	frm1.txtChildItemQty.focus 
	Set gActiveElement = document.activeElement 
	
    Call ggoOper.LockField(Document, "Q")

	If frm1.hBomHistoryFlg.value = "Y" Then
		Call ggoOper.SetReqAttr(frm1.txtEcnNo, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtEcnNo, "Q")
	End If
	
	frm1.cboSupplyType.selectedIndex = 1
	
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG ��																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
						<TD CLASS=TD5 NOWRAP>����</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14xxxU" ALT="����">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="�����"></TD>
						<TD CLASS=TD5 NOWRAP>BOM Type</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBomType" SIZE=6 MAXLENGTH=3 tag="14xxxU" ALT="BOM Type"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>��üǰ</TD>
						<TD CLASS=TD6 NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12xxxU" ALT="��üǰ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="��üǰ"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE CLASS="TB2" CELLSPACING=0>
				<TR>
					<TD WIDTH=100%  valign=top>
						<FIELDSET>
							<TABLE CLASS="TB2" CELLSPACING=0>
								<TR> 
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAcct" SIZE=17 tag="24"  ALT="ǰ�����"></TD>
									<TD CLASS=TD5 NOWRAP>�԰�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE=30 tag="24"  ALT="�԰�"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���ޱ���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProcurType" SIZE=17 tag="24"  ALT="���ޱ���"></TD>
									<TD CLASS=TD5 NOWRAP>�����󱸺�</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboSupplyType" ALT="�����󱸺�" STYLE="Width: 130px;" tag="22"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtChildItemQty CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="��ǰ����ؼ�" tag="22X8Z"> </OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>��ǰ����ش���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtChildUnit" SIZE=8 MAXLENGTH=3 tag="22xxxU"  ALT="��ǰ����ش���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChildUnit" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit('CHILD')"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰ����ؼ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtPrntItemQty CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="��ǰ����ؼ�" tag="22X8Z"> </OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>��ǰ����ش���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPrntUnit" SIZE=8 MAXLENGTH=3 tag="22xxxU"  ALT="��ǰ����ش���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrntUnit" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit('PRNT')"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����L/T</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtSafetyLt CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="����L/T" tag="21X6Z"> </OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>Loss��</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDOUBLESINGLE CLASS=FPDS140 name=txtLossRate CLASSID=<%=gCLSIDFPDS%> SIZE="15" MAXLENGTH="15" ALT="Loss��" tag="21X7Z"> </OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="������"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtValidToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="������"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>���躯���ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtEcnNo" SIZE=18 MAXLENGTH=18 tag="24xxxU" ALT="���躯���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNInfo"></TD>
								</TR>											
								<TR>
									<TD CLASS=TD5 NOWRAP>���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRemark" SIZE=30 MAXLENGTH=1000 tag="21"  ALT="���"></TD>
									<TD CLASS=TD5 NOWRAP>���躯�泻��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtEcnDesc" SIZE=30 MAXLENGTH=50 tag="24"  ALT="���躯�泻��"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>���躯��ٰ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReasonCd" SIZE=6 MAXLENGTH=2 tag="24xxxU"  ALT="���躯��ٰ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonPopup">&nbsp;<INPUT TYPE=TEXT NAME="txtReasonNm" SIZE=20 tag="24"  ALT="����ٰŸ�"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								</TR>

							</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD HEIGHT=30>
			<TABLE CLASS="basicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT>
					<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBomType" tag="24">
<INPUT TYPE=HIDDEN NAME="hBomHistoryFlg" tag="24" value="Y">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
