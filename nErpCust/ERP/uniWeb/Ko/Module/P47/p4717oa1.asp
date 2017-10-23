<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : production
'*  2. Function Name        : 
'*  3. Program ID           : p4717oa1
'*  4. Program Name         : (p)�ڿ��Һ����(�ڿ���) 
'*  5. Program Desc         :  P4717OA1.EBR (�ڿ���) ��� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001.12.12
'*  9. Modifier (First)     : Jaehyun Chen
'* 10. Modifier (Last)      : Jaehyun Chen
'* 11. Comment 
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin             :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'#########################################################################################################-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--��:Print Program needs this vbs file-->
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop
'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
'----------------  ���� Global ������ ����  -----------------------------------------------------------
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim startDate
Dim EndDate
Dim strYear
Dim strMonth
Dim strDay
Dim BaseDate

BaseDate = "<%=GetsvrDate%>"
Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)	    	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")	   '��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

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
    IsOpenPop = False     
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
	frm1.txtStartDt.Text =	StartDate
	frm1.txtEndDt.Text = EndDate
	
End Sub

'========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================
Sub InitComboBox()
<%

%>
        
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "P","NOCOOKIE","OA") %>
End Sub

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
    
	Call ggoOper.FormatField(Document, "x",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitComboBox	
	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables
   	
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtFromResourceCd.focus 
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

Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)									'��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()                                         '��: Protect system from crashing
    Call parent.FncPrint()
End Function

'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
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
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================

Function BtnPrint()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	
	dim strUrl
	dim arrParam, arrField, arrHeader
	Dim strEbrFile
	Dim objName
	
	If frm1.txtFromResourceCd.value = "" Then
		frm1.txtFromResourceNm.value = "" 
	End If
	
	If frm1.txtToResourceCd.value = "" Then
		frm1.txtToResourceNm.value = "" 
	End If
	
	If frm1.txtFromResourceGroup.value = "" Then
		frm1.txtFromResourceGroupNm.value = "" 
	End If	
	
	If frm1.txtToResourceGroup.value = "" Then
		frm1.txtToResourceGroupNm.value = "" 
	End If	
	
    Call BtnDisabled(1)	
	
	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))	
	
	var2 = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	var3 = UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If Trim(frm1.txtFromResourceCd.value) = "" Then
		var4 = "0"
	Else
		var4 = Trim(frm1.txtFromResourceCd.value)
	End If
	
	If Trim(frm1.txtToResourceCd.value) = "" Then
		var5 = "zzzzzzzzzzz"
	Else
		var5 = Trim(frm1.txtToResourceCd.value)
	End If
	
	If Trim(frm1.txtFromResourceGroup.value) = "" Then
		var6 = "0"
	Else
		var6 = Trim(frm1.txtFromResourceGroup.value)  
	End If
	
	If Trim(frm1.txtToResourceGroup.value) = "" Then
		var7 = "zzzzzzzzzzz"
	Else
		var7 = Trim(frm1.txtToResourceGroup.value)
	End If

	strUrl = strUrl & "plant_cd|" & var1 
	strUrl = strUrl & "|from_consumed_dt|" & var2 
	strUrl = strUrl & "|to_consumed_dt|" & var3
	strUrl = strUrl & "|fr_resource_cd|" & var4 
	strUrl = strUrl & "|to_resource_cd|" & var5 
	strUrl = strUrl & "|fr_resource_group|" & var6 
	strUrl = strUrl & "|to_resource_group|" & var7 

'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
'----------------------------------------------------------------
	strEbrFile = "p4717oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	

	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement
	
End Function


'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function BtnPreview() 
'On Error Resume Next                                                    '��: Protect system from crashing
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	dim var7
	
	dim strUrl
	dim arrParam, arrField, arrHeader
	Dim strEbrFile
	Dim objName
	
	Call BtnDisabled(1)
	
	If frm1.txtFromResourceCd.value = "" Then
		frm1.txtFromResourceNm.value = "" 
	End If
	
	If frm1.txtToResourceCd.value = "" Then
		frm1.txtToResourceNm.value = "" 
	End If
	
	If frm1.txtFromResourceGroup.value = "" Then
		frm1.txtFromResourceGroupNm.value = "" 
	End If	
	
	If frm1.txtToResourceGroup.value = "" Then
		frm1.txtToResourceGroupNm.value = "" 
	End If	
	
	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))	
	
	var2 = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	var3 = UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If Trim(frm1.txtFromResourceCd.value) = "" Then
		var4 = "0"
	Else
		var4 = Trim(frm1.txtFromResourceCd.value  )
	End If
	
	If Trim(frm1.txtToResourceCd.value) = "" Then
		var5 = "zzzzzzzzzzz"
	Else
		var5 = Trim(frm1.txtToResourceCd.value)
	End If
	
	If Trim(frm1.txtFromResourceGroup.value) = "" Then
		var6 = "0"
	Else
		var6 = Trim(frm1.txtFromResourceGroup.value  )
	End If
	
	If Trim(frm1.txtToResourceGroup.value) = "" Then
		var7 = "zzzzzzzzzzz"
	Else
		var7 = Trim(frm1.txtToResourceGroup.value)
	End If
	
	strUrl = strUrl & "plant_cd|" & var1 
	strUrl = strUrl & "|from_consumed_dt|" & var2 
	strUrl = strUrl & "|to_consumed_dt|" & var3
	strUrl = strUrl & "|fr_resource_cd|" & var4 
	strUrl = strUrl & "|to_resource_cd|" & var5 
	strUrl = strUrl & "|fr_resource_group|" & var6 
	strUrl = strUrl & "|to_resource_group|" & var7 

	'call FncEBRPreview("p4717oa1.ebr", strUrl)
	strEbrFile = "p4717oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRPreview(objName, strUrl)
		
	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function


'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
	On Error Resume Next        
End Function



'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------

Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"		' �˾� ��Ī 
	arrParam(1) = "B_PLANT"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "����"			' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"			' Field��(0)
    arrField(1) = "PLANT_NM"			' Field��(1)
    
    arrHeader(0) = "����"			' Header��(0)
    arrHeader(1) = "�����"			' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenFromResource()  -------------------------------------------------
'	Name : OpenFromResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtFromResourceCd.Value)
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
		Call SetFromResource(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtFromResourceCd.focus
		
End Function

'------------------------------------------  OpenToResource()  -------------------------------------------------
'	Name : OpenToResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenToResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	IsOpenPop = True
	
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtToResourceCd.Value)
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
		Call SetToResource(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtToResourceCd.focus
		
End Function
'------------------------------------------  OpenFromResourceGroup()  -------------------------------------------------
'	Name : OpenFromResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtFromResourceGroup.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�ڿ��׷��˾�"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtFromResourceGroup.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
				  			
	arrParam(5) = "�ڿ��׷�"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ��׷�"		
    arrHeader(1) = "�ڿ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetFromResourceGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtFromResourceGroup.focus
	
End Function

'------------------------------------------  OpenToResourceGroup()  -------------------------------------------------
'	Name : OpenToResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenToResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If UCase(frm1.txtToResourceGroup.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "�ڿ��׷��˾�"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtToResourceGroup.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
				  			
	arrParam(5) = "�ڿ��׷�"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ��׷�"		
    arrHeader(1) = "�ڿ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetToResourceGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtToResourceGroup.focus
	
End Function

Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function
Function SetFromResource(ByVal arrRet)
	frm1.txtFromResourceCd.value = arrRet(0) 
	frm1.txtFromResourceNm.value = arrRet(1)
End Function

Function SetToResource(ByVal arrRet)
	frm1.txtToResourceCd.value = arrRet(0) 
	frm1.txtToResourceNm.value = arrRet(1) 
End Function
Function SetFromResourceGroup(ByVal arrRet)
	frm1.txtFromResourceGroup.value = arrRet(0)
	frm1.txtFromResourceGroupNm.value = arrRet(1)  
End Function
Function SetToResourceGroup(ByVal arrRet)
	frm1.txtToResourceGroup.value = arrRet(0)
	frm1.txtToResourceGroupNm.value = arrRet(1)  
End Function

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtEndDt.Focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ڿ��Һ����(�ڿ���)</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" colspan="2">
			<TABLE CLASS="BasicTB" CELLSPACING=0 >	
	    		<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="x2xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�����">
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>�ڿ��ڵ�</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromResourceCd" SIZE=10 MAXLENGTH=10 tag="x1xxxU" ALT="�ڿ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromResourceCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromResourceNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�ڿ���">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToResourceCd" SIZE=10 MAXLENGTH=10 tag="x1xxxU" ALT="�ڿ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToResourceCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtToResourceNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�ڿ���">&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>�ڿ��׷�</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromResourceGroup" SIZE=10 MAXLENGTH=10 tag="x1xxxU" ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromResourceGroup" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromResourceGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromResourceGroupNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�ڿ��׷��">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToResourceGroup" SIZE=10 MAXLENGTH=10 tag="x1xxxU" ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToResourceGroup" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToResourceGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtToResourceGroupNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�ڿ��׷��">&nbsp;
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ڿ��Һ���</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4717oa1_I567901867_txtStartDt.js'></script>
										&nbsp;~&nbsp; 
										<script language =javascript src='./js/p4717oa1_I769342577_txtEndDt.js'></script>								
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->

</BODY>
</HTML>
