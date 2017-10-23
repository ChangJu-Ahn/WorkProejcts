
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  p1501ma1.asp
'*  4. Program Name         :  Resource Management
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/04
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_QRY_ID = "p1501mb1.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "p1501mb2.asp"											'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_DEL_ID = "p1501mb3.asp"	
Const BIZ_PGM_LOOKUP_ID = "p1501mb4.asp"
Const BIZ_PGM_LOOKUP_CUR_ID = "p1502mb4.asp"	'�߰� 2003-04-17

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo
Dim iDBSYSDate
Dim StartDate, EndDate
Dim IsOpenPop          
Dim lgRdoOldVal
Dim lgCurCd
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
    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
End Sub

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
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
	iDBSYSDate = "<%=GetSvrDate%>"											'��: DB�� ���� ��¥�� �޾ƿͼ� ���۳�¥�� ����Ѵ�.
	StartDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, gDateFormat)
	frm1.txtValidFromDt.text  = startdate
	frm1.txtValidToDt.text	= UniConvYYYYMMDDToDate(gDateFormat, "2999","12","31")
	frm1.cboResourceType.value = "L"
	frm1.txtNoOfResource.Value = "1"
	frm1.txtCostType.value = "E"
	frm1.txtEfficiency.text = "100"
	frm1.txtUtilization.text = "100"
	frm1.rdoRunCrp2.checked = true
	frm1.rdoRunRccp2.checked = true
End Sub

Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1502", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboResourceType, lgF0, lgF1, Chr(11))
    Call CommonQueryRs(" RULE_TYPE,DESCRIPTION "," P_APS_RULE_DETAIL "," RULE_TYPE_CD = " & FilterVar("RSSLRL", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    '<!--	RCCP ���� ���� Start
    'lgF0 = "0" & Chr(11) & lgF0
    'lgF1 = "" & Chr(11) & lgF1
    'Call SetCombo2(frm1.txtSelectionRule, lgF0, lgF1, Chr(11))
    'RCCP ���� ���� End	-->
    Call CommonQueryRs(" RULE_TYPE,DESCRIPTION "," P_APS_RULE_DETAIL "," RULE_TYPE_CD = " & FilterVar("RSSQRL", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = "0" & Chr(11) & lgF0
    lgF1 = "" & Chr(11) & lgF1
    '<!--	RCCP ���� ���� Start
    'Call SetCombo2(frm1.txtSequenceRule, lgF0, lgF1, Chr(11))
    'RCCP ���� ���� End	-->
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
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    arrField(2) = "CUR_CD"
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    arrHeader(2) = "��ȭ�ڵ�"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()
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
	arrParam(2) = Trim(frm1.txtResourceCd1.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
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
	
	Call SetFocusToDocument("M")
	frm1.txtResourceCd1.focus
	
End Function

'------------------------------------------  OpenResourceGroup()  -------------------------------------------------
'	Name : OpenResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtResourceGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�ڿ��׷��˾�"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtResourceGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
	arrParam(5) = "�ڿ��׷�"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ��׷�"		
    arrHeader(1) = "�ڿ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetResourceGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtResourceGroupCd.focus
	
End Function

'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	
	If IsOpenPop = True Then
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "�ڿ����ش����˾�"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtResourceUnitCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DIMENSION = " & FilterVar("TM", "''", "S") & ""			
	arrParam(5) = "�ڿ����ش���"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "�ڿ����ش���"		
    arrHeader(1) = "�ڿ����ش�����"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceUnitCd.focus
		
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtCurCd.value      = UCase(arrRet(2))
	lgCurCd = UCase(arrRet(2))		
End Function

'------------------------------------------  SetResource()  --------------------------------------------------
'	Name : SetResource()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResource(byval arrRet)
	frm1.txtResourceCd1.Value    = arrRet(0)		
	frm1.txtResourceNm1.Value    = arrRet(1)		
End Function

'------------------------------------------  SetResourceGroup()  --------------------------------------------------
'	Name : SetResourceGroup()
'	Description : ResourceGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetResourceGroup(byval arrRet)
	frm1.txtResourceGroupCd.Value    = arrRet(0)		
	frm1.txtResourceGroupNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : Resource Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(byval arrRet)
	frm1.txtResourceUnitCd.Value    = arrRet(0)	
	frm1.txtResourceUnitCd1.value   = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function ChkValidData()
	ChkValidData = False
	
	With frm1
		If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function       
		'<!--	RCCP ���� ���� Start
		'If frm1.txtSelectionRule.value <> "" Then
		'	If CInt(frm1.txtSelectionRule.value) > 39 Then
		'		Call DisplayMsgBox("970025","X", "���ñ�Ģ","20")
		'		.txtSelectionRule.focus 
		'		Set gActiveElement = document.activeElement 
		'		Exit Function
		'	End IF	
		'End If
		
		'If frm1.txtSequenceRule.value <> "" Then
		'	If CInt(frm1.txtSequenceRule.value) > 39 Then
		'		Call DisplayMsgBox("970025","X", "������Ģ","20")
		'		.txtSequenceRule.focus 
		'		Set gActiveElement = document.activeElement 
		'		Exit Function
		'	End IF	
		'End If
		'RCCP ���� ���� End	-->
	End With
	
	ChkValidData = True
End Function

Sub ChkNumKeyPress()
	Dim KeyCode
	KeyCode = window.event.keyCode 
	
	If KeyCode < 48 Or KeyCode > 57 Then
		window.event.keyCode = 8
		Exit Sub
	End If
End Sub

'==========================================  2.5.6 LookUpRuleType() =======================================
'=	Event Name : LookUpRuleType																				=
'=	Event Desc :																						=
'========================================================================================================
Sub LookUpRuleType()
	LayerShowHide(1) 
		
	Dim strVal

	strVal = BIZ_PGM_LOOKUP_ID & "?txtMode=" & parent.UID_M0001				'��: �����Ͻ� ó�� ASP�� ���� 	
	'<!--	RCCP ���� ���� Start
	'strVal = strVal & "&txtSelectionRule=" & Trim(frm1.txtSelectionRule.value)		'��: ��ȸ ���� ����Ÿ 
	'RCCP ���� ���� End	-->
	strVal = strVal & "&PrevNextFlg=" & ""	
	strVal = strVal & "&lgCurDate=" & startdate
	
	Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
End Sub

'==========================================  2.5.7 LookUpItemOk() =======================================
'=	Event Name : LookUpItemOk																				=
'=	Event Desc :																						=
'========================================================================================================
Sub LookUpRuleTypeOk()
	IsOpenPop = False
End Sub

Sub LookUpRuleTypeNotOk()
	IsOpenPop = False
End Sub

Function CurCdLookUp()
		Dim strVal
		lgCurCd = ""
		frm1.txtCurCd.value = ""
		
		strVal = BIZ_PGM_LOOKUP_CUR_ID & "?txtMode=" & parent.UID_M0001					'��: �����Ͻ� ó�� ASP�� ���� 	
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)		'��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&PrevNextFlg=" & ""	
	
		Call RunMyBizASP(MyBizASP, strVal)									'��: �����Ͻ� ASP �� ���� 
End Function

Function CurCdLooKUpOk()
		lgCurCd = frm1.txtCurCd.value 
		IsOpenPop = False
End Function

Function CurCdLooKUpNotOk()
		
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
	Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","6","0")
	Call AppendNumberPlace("7","3","2")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    '----------  Coding part  -------------------------------------------------------------
    Call InitComboBox
    Call SetDefaultVal
    Call InitVariables
    Call SetToolbar("11101000000011")
    If parent.gPlant <> "" then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		Call CurCdLooKUp()
		frm1.txtResourceCd1.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

Sub txtResourceEa_Change()
    lgBlnFlgChgValue = True    
	frm1.txtResourceEa1.value = frm1.txtResourceEa.value	
End Sub

Sub txtMfgCost_Change()
    lgBlnFlgChgValue = True    	
End Sub

Sub txtResourceUnitCd_OnChange()
	frm1.txtResourceUnitCd1.value = frm1.txtResourceUnitCd.value 
End Sub

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
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
'   Event Name : txtValidFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
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
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtValidToDt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtNoOfResource_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtEfficiency_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtUtilization_Change()
    lgBlnFlgChgValue = True
End Sub
Sub txtOverloadTol_Change()
    lgBlnFlgChgValue = True
End Sub

Sub rdoRunRccp1_onChange()
    lgBlnFlgChgValue = True
End Sub
Sub rdoRunRccp2_onChange()
    lgBlnFlgChgValue = True
End Sub
Sub rdoRunCrp1_onChange()
    lgBlnFlgChgValue = True
End Sub
Sub rdoRunCrp2_onChange()
    lgBlnFlgChgValue = True
End Sub

Sub rdoInfiniteResourceFlg1_OnClick()
	If lgRdoOldVal = 1 Then Exit Sub
	
	lgBlnFlgChgValue = True
	lgRdoOldVal = 1
End Sub

Sub cboResourceType_onchange()
	If frm1.cboResourceType.value = "M" Then
		frm1.txtCostType.value = "E"
	Else
		frm1.txtCostType.value = "L"
	End IF
	lgBlnFlgChgValue = True
End Sub

'<!--	RCCP ���� ���� Start
'Sub txtSelectionRule_onKeyPress()
'	Call ChkNumKeyPress
'	lgBlnFlgChgValue = True
'End Sub
'
'Sub txtSequenceRule_onKeyPress()
'	Call ChkNumKeyPress
'	lgBlnFlgChgValue = True
'End Sub
'RCCP ���� ���� End	-->

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'----------  Coding part  ------------------------------------------------------------- 


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
    
    FncQuery = False															'��: Processing is NG
    
    Err.Clear																	'��: Protect system from crashing

   '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
 '-----------------------
    'Erase contents area
    '----------------------- 
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
		
	If frm1.txtResourceCd1.value = "" Then
		frm1.txtResourceNm1.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables

 '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
 '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then   
		Exit Function           
    End If     										'��: Query db data
       
    FncQuery = True																'��: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False																'��: Processing is NG
    
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")					'��: "����Ÿ�� ����Ǿ����ϴ�. �ű��Է��� �Ͻðڽ��ϱ�?"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    frm1.txtResourceCd1.value = ""
    frm1.txtResourceNm1.value = "" 
    Call ggoOper.ClearField(Document, "2")                                      '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
	Call SetToolbar("11101000000011")    
	frm1.txtResourceCd2.focus 
	Set gActiveElement = document.activeElement 
	frm1.txtCurCd.value = lgCurCd	
    FncNew = True																'��: Processing is OK
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    Dim intRetCD
    
    FncDelete = False														'��: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    If DbDelete = False Then   
		Exit Function           
    End If     														'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
    If ValidDateCheck(frm1.txtValidFromDt, frm1.txtValidToDt) = False Then Exit Function
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          '��: No data changed!!
        Exit Function
    End If
	
    If Not chkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
	
	If UniCDbl(frm1.txtResourceEa.Text) <= CDbl(0) Then
		Call DisplayMsgBox("970022","X",frm1.txtResourceEa.alt,"0")
		frm1.txtResourceEa.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
    If DbSave = False Then
		Exit Function           
    End If     				                                                '��: Save db data

    FncSave = True                                                          '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE												'��: Indicates that current mode is Crate mode
    
    Call ggoOper.LockField(Document, "N")									'��: This function lock the suitable field
	Call SetToolbar("11101000000011")
	
    frm1.txtResourceCd1.value = ""
    frm1.txtResourceNm1.value = ""
    
    frm1.txtResourceCd2.value = ""
	
	frm1.txtValidFromDt.text  = startdate
	frm1.txtValidToDt.text	= UniConvYYYYMMDDToDate(gDateFormat, "2999","12","31")
    
    frm1.txtResourceCd2.focus
    Set gActiveElement = document.activeElement 
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
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    Dim strVal
    Dim IntRetCD
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                               '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables

    Err.Clear                                                               '��: Protect system from crashing
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtResourceCd1=" & Trim(frm1.txtResourceCd1.value)				'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & "P"
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    Dim strVal
	Dim IntRetCD
	
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                               '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
    
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")					'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
    
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call SetDefaultVal
    Call InitVariables															'��: Initializes local global variables
    
    Err.Clear                                                               '��: Protect system from crashing
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)			'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtResourceCd1=" & Trim(frm1.txtResourceCd1.value)				'��: ��ȸ ���� ����Ÿ 
	strVal = strVal & "&txtUpdtUserId=" & parent.gUsrID
	strVal = strVal & "&PrevNextFlg=" & "N"
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)											'��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")				'��: "Will you destory previous data"
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
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1) 

    strVal = BIZ_PGM_DEL_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
	strVal = strVal & "&txtResourceCd=" & Trim(frm1.hResourceCd.value)
	
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	Call InitVariables
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Err.Clear                                                               '��: Protect system from crashing
    DbQuery = False                                                         '��: Processing is NG
    
    Dim strVal
	
	LayerShowHide(1) 
		
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtResourceCd1=" & Trim(frm1.txtResourceCd1.value)
    strVal = strVal & "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)
	
    DbQuery = True
End Function

'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=========================================================================================================
Function DbQueryOk()
    frm1.hPlantCd.value = frm1.txtPlantCd.value 
    lgCurCd = frm1.txtCurCd.value
	
    lgIntFlgMode = parent.OPMD_UMODE
    lgBlnFlgChgValue = false
        
    Call ggoOper.LockField(Document, "Q")

	Call SetToolbar("11111000111111")
	
	frm1.txtResourceNm2.focus
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 
	Dim BlnRetCd

    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG

    Dim strVal

	BlnRetCd = ChkValidData

	If BlnRetCd = False Then
		Exit Function
	End if

	LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										
	End With

    DbSave = True                                                           '��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()															'��: ���� ������ ���� ���� 
    frm1.txtResourceCd1.value = frm1.txtResourceCd2.value 
    frm1.txtResourceNm1.value = frm1.txtResourceNm2.value 

    Call InitVariables
    
    Call MainQuery()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ����</font></td>
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
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()"> <INPUT TYPE=TEXT ID="txtPlantNm" NAME="arrCond" SIZE=50 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
									<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd1" SIZE=20 MAXLENGTH=10 tag="12XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm1" SIZE=50 tag="14"></TD>
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
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=100% valign=top>
									<FIELDSET>
										<LEGEND>�Ϲ�����</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>�ڿ�</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd2" SIZE=20 MAXLENGTH=10 tag="23XXXU" ALT="�ڿ�">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm2" SIZE=50 MAXLENGTH=40 tag="22XXXX" ALT="�ڿ���"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�ڿ��׷�</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroupCd" SIZE=20 MAXLENGTH=10 tag="23XXXU" ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResourceGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroupNm" SIZE=50 tag="24"></TD>
											</TR>	
											<TR>
												<TD CLASS=TD5 NOWRAP>�ڿ�����</TD>
												<TD CLASS=TD656 NOWRAP><SELECT NAME="cboResourceType" ALT="�ڿ�����" STYLE="Width: 98px;" tag="22"></SELECT></TD>																								
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�ڿ���</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I434898900_txtNoOfResource.js'></script>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>ȿ��</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I599973791_txtEfficiency.js'></script>&nbsp;%
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>������</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I784308445_txtUtilization.js'></script>&nbsp;%
												</TD>
											</TR>		
											<TR ID=Q1>
												<TD CLASS=TD5 NOWRAP>RCCP���ϰ����</TD>
												<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="2X" ID="rdoRunRccp1" VALUE="Y"><LABEL FOR="rdoRunRccp1">��</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunRccp" TAG="2X" ID="rdoRunRccp2" VALUE="N"><LABEL FOR="rdoRunRccp2">�ƴϿ�</LABEL></TD>
											</TR>
											<TR ID=Q2>
												<TD CLASS=TD5 NOWRAP>CRP���ϰ����</TD>
												<TD CLASS=TD6 NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="24" ID="rdoRunCrp1" VALUE="Y"><LABEL FOR="rdoRunCrp1">��</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoRunCrp" TAG="24" ID="rdoRunCrp2" VALUE="N"><LABEL FOR="rdoRunCrp2">�ƴϿ�</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�����������</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I161376506_txtOverloadTol.js'></script>&nbsp;%
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>�ڿ����ؼ���</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I833706259_txtResourceEa.js'></script>												
												</TD>
											</TR>																																	
											<TR>
												<TD CLASS=TD5 NOWRAP>�ڿ����ش���</TD>
												<TD CLASS=TD656 NOWRAP><INPUT TYPE=TEXT NAME="txtResourceUnitCd" SIZE=5 MAXLENGTH=3 tag="22XXXU" ALT="�ڿ����ش���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnResourceGroupCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>���ش����� �����������</TD>
												<TD CLASS=TD656 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>															
																<script language =javascript src='./js/p1501ma1_I324865192_txtMfgCost.js'></script>
															</TD>
															<TD>											
																&nbsp;<INPUT TYPE=TEXT NAME="txtCurCd" tag=24 SIZE=5 MAXLENGTH=3 ALT="��ȭ�ڵ�">&nbsp;/&nbsp;
															</TD>
															<TD>
																<script language =javascript src='./js/p1501ma1_I870693814_txtResourceEa1.js'></script>												
															</TD>
															<TD>
																&nbsp;<INPUT TYPE=TEXT NAME="txtResourceUnitCd1" SIZE=5 MAXLENGTH=3 tag="24XXXU" ALT="�ڿ����ش���">
															</TD>
														</TR>
													</TABLE>												
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>��ȿ�Ⱓ</TD>
												<TD CLASS=TD656 NOWRAP>
													<script language =javascript src='./js/p1501ma1_I130084407_txtValidFromDt.js'></script>
													&nbsp;~&nbsp;
													<script language =javascript src='./js/p1501ma1_I571680456_txtValidToDt.js'></script>										
												</TD>
											</TR>
											</TABLE>
									</FIELDSET>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hResourceCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtCostType" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
