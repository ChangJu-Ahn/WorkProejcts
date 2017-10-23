<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : p1410oa1.asp
'*  4. Program Name         : ECN������� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003-03-10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Woo Guen
'* 10. Modifier (Last)      : 
'* 11. Comment              :
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = VBSCRIPT>

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
<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop
 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim CurBaseDate

CurBaseDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

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
	frm1.txtECNBaseDt.Text = CurBaseDate
End Sub
 '========================== 2.2.6 InitComboBox()  =====================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================= 
Sub InitComboBox()

    On Error Resume Next
    Err.Clear
    
    'Call SetCombo(frm1.cboECNStatus, "1", "Active")
	'Call SetCombo(frm1.cboECNStatus, "2", "Inactive")	
	
    'Call SetCombo(frm1.cboEBOM, "Y", "��")
	'Call SetCombo(frm1.cboEBOM, "N", "�ƴϿ�")
	
    'Call SetCombo(frm1.cboMBOM, "Y", "��")
	'Call SetCombo(frm1.cboMBOM, "N", "�ƴϿ�")	
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
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
    
	Call ggoOper.FormatField(Document, "X",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
   
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitComboBox	
	Call SetDefaultVal
    Call InitVariables		'��: Initializes local global variables
    	
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 

	frm1.txtFromECNCd.focus 
	Set gActiveElement = document.activeElement
   
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
	Dim var6
	Dim var7
	Dim var8
	Dim var9
	Dim var10
	Dim var11
	
	Dim strUrl, strEbrFile, objName
    
    Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
    
	If frm1.txtFromECNCd.value = "" Then
		frm1.txtFromECNDesc.value = "" 
	End If	
	
	If frm1.txtToECNCd.value = "" Then
		frm1.txtToECNDesc.value = "" 
	End If
	
	If frm1.txtReasonCd.value = "" Then
		frm1.txtReasonDesc.value = "" 
	End If	
	
    
	var1 = UCase(Trim(frm1.txtFromECNCd.value))
	var2 = UCase(Trim(frm1.txtToECNCd.value))
	'var3 = UNIConvDate(frm1.txtECNBaseDt.Text)
	var3 = UniConvDateAtoB(frm1.txtECNBaseDt.Text,parent.gDateFormat,parent.gServerDateFormat)

	If frm1.txtReasonCd.value = "" Then
		var4 = "0"
		var5 = "ZZ"
	Else
		var4 = Trim(frm1.txtReasonCd.value)
		var5 = Trim(frm1.txtReasonCd.value)
	End If
		
	If frm1.cboECNStatus2.checked  = True Then	  
		var6 = "1"									
		var7 = "1"											 
	ElseIf frm1.cboECNStatus3.checked  = True Then
		var6 = "2"	
		var7 = "2"								
	Else
		var6 = "0"	
		var7 = "2"											 
	End If
	
	If frm1.cboEBOM2.checked  = True Then	  
		var8 = "Y"									
		var9 = "Y"											 
	ElseIf frm1.cboEBOM3.checked  = True Then
		var8 = "N"	
		var9 = "N"								
	Else
		var8 = "0"	
		var9 = "Y"											 
	End If

	If frm1.cboMBOM2.checked  = True Then	  
		var10 = "Y"									
		var11 = "Y"											 
	ElseIf frm1.cboMBOM3.checked  = True Then
		var10 = "N"	
		var11 = "N"								
	Else
		var10 = "0"	
		var11 = "Y"											 
	End If
	
	strUrl = strUrl & "fr_ecn_no|"		& var1 
	strUrl = strUrl & "|to_ecn_no|"		& var2 
	strUrl = strUrl & "|ecn_base_dt|"	& var3
	strUrl = strUrl & "|fr_reason_cd|"	& var4
	strUrl = strUrl & "|to_reason_cd|"	& var5 
	strUrl = strUrl & "|fr_ecn_status|" & var6
	strUrl = strUrl & "|to_ecn_status|" & var7 
	strUrl = strUrl & "|fr_ebom_flg|"	& var8 
	strUrl = strUrl & "|to_ebom_flg|"	& var9 
	strUrl = strUrl & "|fr_mbom_flg|"	& var10
	strUrl = strUrl & "|to_mbom_flg|"	& var11 
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	strEbrFile = "p1410oa1"
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
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	Dim var9
	Dim var10
	Dim var11
	
	Dim strUrl, strEbrFile, objName
	
	Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
    
	If frm1.txtFromECNCd.value = "" Then
		frm1.txtFromECNDesc.value = "" 
	End If	
	
	If frm1.txtToECNCd.value = "" Then
		frm1.txtToECNDesc.value = "" 
	End If
	
	If frm1.txtReasonCd.value = "" Then
		frm1.txtReasonDesc.value = "" 
	End If	
    

	var1 = UCase(Trim(frm1.txtFromECNCd.value))
	var2 = UCase(Trim(frm1.txtToECNCd.value))
	'var3 = UNIConvDate(frm1.txtECNBaseDt.Text)
	var3 = UniConvDateAtoB(frm1.txtECNBaseDt.Text,parent.gDateFormat,parent.gServerDateFormat)

	If frm1.txtReasonCd.value = "" Then
		var4 = "0"
		var5 = "ZZ"
	Else
		var4 = Trim(frm1.txtReasonCd.value)
		var5 = Trim(frm1.txtReasonCd.value)
	End If
			
	If frm1.cboECNStatus2.checked  = True Then	  
		var6 = "1"									
		var7 = "1"											 
	ElseIf frm1.cboECNStatus3.checked  = True Then
		var6 = "2"	
		var7 = "2"								
	Else
		var6 = "0"	
		var7 = "2"											 
	End If
	
	If frm1.cboEBOM2.checked  = True Then	  
		var8 = "Y"									
		var9 = "Y"											 
	ElseIf frm1.cboEBOM3.checked  = True Then
		var8 = "N"	
		var9 = "N"								
	Else
		var8 = "0"	
		var9 = "Y"											 
	End If

	If frm1.cboMBOM2.checked  = True Then	  
		var10 = "Y"									
		var11 = "Y"											 
	ElseIf frm1.cboMBOM3.checked  = True Then
		var10 = "N"	
		var11 = "N"								
	Else
		var10 = "0"	
		var11 = "Y"											 
	End If

	strUrl = strUrl & "fr_ecn_no|"		& var1 
	strUrl = strUrl & "|to_ecn_no|"		& var2 
	strUrl = strUrl & "|ecn_base_dt|"	& var3
	strUrl = strUrl & "|fr_reason_cd|"	& var4
	strUrl = strUrl & "|to_reason_cd|"	& var5 
	strUrl = strUrl & "|fr_ecn_status|" & var6
	strUrl = strUrl & "|to_ecn_status|" & var7 
	strUrl = strUrl & "|fr_ebom_flg|"	& var8 
	strUrl = strUrl & "|to_ebom_flg|"	& var9 
	strUrl = strUrl & "|fr_mbom_flg|"	& var10
	strUrl = strUrl & "|to_mbom_flg|"	& var11 

'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	strEbrFile = "p1410oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRPreview(objName, strUrl)
'----------------------------------------------------------------	

	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'===========================================  2.4.1 POP-UP Open �Լ�()  =================================
'=	Name : Open???()																					=
'=	Description : POP-UP Open																			=
'========================================================================================================

'------------------------------------------  OpenFromECNCd()  ----------------------------------------------
'	Name : OpenFromECNCd()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenFromECNCd()

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtFromECNCD.value)	' ECNNo
	arrParam(1) = ""							' ReasonCd
	arrParam(2) = ""							' Status
	arrParam(3) = ""							' EBomFlg
	arrParam(4) = ""							' MBomFlg

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetFromECNCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtFromECNCD.Focus
	
End Function

'------------------------------------------  OpenToECNCd()  ----------------------------------------------
'	Name : OpenToECNCd()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenToECNCd()

	Dim arrRet
	Dim arrParam(4), arrField(10)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtToECNCD.value)	' ECNNo
	arrParam(1) = ""							' ReasonCd
	arrParam(2) = ""							' Status
	arrParam(3) = ""							' EBomFlg
	arrParam(4) = ""							' MBomFlg

	iCalledAspName = AskPRAspName("P1410PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P1410PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
 
	If arrRet(0) <> "" Then
		Call SetToECNCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtToECNCD.Focus
	
End Function

'------------------------------------------  OpenReasonCd()  ------------------------------------------
'	Name : OpenReasonCd()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenReasonCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
   
	If IsOpenPop = True Then 
		IsOpenPop = False
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
	
    arrField(0) = "MINOR_CD"							' Field��(0)
    arrField(1) = "MINOR_NM"							' Field��(1)
        
    arrHeader(0) = "���躯��ٰ�"					' Header��(0)
    arrHeader(1) = "���躯��ٰŸ�"					' Header��(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetReasonCd(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	Frm1.txtReasonCd.Focus	
	
End Function


'=======================================  2.4.2 POP-UP Return�� ���� �Լ�  ==============================
'=	Name : Set???()																						=
'=	Description : Reference �� POP-UP�� Return���� �޴� �κ�											=
'========================================================================================================

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetFromECNCd()  ------------------------------------------------
'	Name : SetFromECNCd()
'	Description : ECN Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetFromECNCd(byval arrRet)
	frm1.txtFromECNCd.Value    = arrRet(0)		
	frm1.txtFromECNDesc.Value  = arrRet(1)
	
	frm1.txtToECNCd.focus
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetToECNCd()  ------------------------------------------------
'	Name : SetToECNCd()
'	Description : ECN Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetToECNCd(byval arrRet)
	frm1.txtToECNCd.Value    = arrRet(0)		
	frm1.txtToECNDesc.Value  = arrRet(1)
	
	frm1.txtReasonCd.focus
	Set gActiveElement = document.activeElement 
End Function

'------------------------------------------  SetReasonCd()  --------------------------------------------------
'	Name : SetReasonCd()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function SetReasonCd(byval arrRet)
	frm1.txtReasonCd.Value	= arrRet(0)
	frm1.txtReasonDesc.Value  = arrRet(1)	
		
	frm1.txtReasonCd.focus
	Set gActiveElement = document.activeElement
End Function


'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtECNBaseDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtECNBaseDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtECNBaseDt.Focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

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
		<TD HEIGHT=5 colspan="2">&nbsp;<% ' ���� ���� %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���躯���������</font></td>
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
								    <TD CLASS="TD5" NOWRAP>���躯���ȣ</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtFromECNCd" SIZE=20 MAXLENGTH=18 tag="X2XXXU" ALT="���躯���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromECNCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromECNCd()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtFromECNDesc" SIZE=40 tag="X4" ALT="���躯�泻��">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtToECNCd" SIZE=20 MAXLENGTH=18 tag="X2XXXU" ALT="���躯���ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToECNCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToECNCd()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtToECNDesc" SIZE=40 tag="X4" ALT="���躯�泻��">&nbsp;
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/p1410oa1_I799672675_txtECNBaseDt.js'></script>							
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>���躯��ٰ�</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtReasonCd" SIZE=12 MAXLENGTH=2 tag="X1XXXU" ALT="���躯��ٰ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReasonCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonCd()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtReasonDesc" SIZE=40 tag="X4" ALT="���躯��ٰŸ�">&nbsp;
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���躯�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboECNStatus" tag="1X" CHECKED ID="cboECNStatus1" VALUE="A"><LABEL FOR="cboECNStatus1">��ü</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboECNStatus" tag="1X" ID="cboECNStatus2" VALUE="Y"><LABEL FOR="cboECNStatus2">Active</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboECNStatus" tag="1X" ID="cboECNStatus3" VALUE="N"><LABEL FOR="cboECNStatus3">Inactive</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����BOM �ݿ�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBOM" tag="1X" CHECKED ID="cboEBOM1" VALUE="A"><LABEL FOR="cboEBOM1">��ü</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBOM" tag="1X" ID="cboEBOM2" VALUE="Y"><LABEL FOR="cboEBOM2">��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboEBOM" tag="1X" ID="cboEBOM3" VALUE="N"><LABEL FOR="cboEBOM3">�ƴϿ�</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����BOM �ݿ�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBOM" tag="1X" CHECKED ID="cboMBOM1" VALUE="A"><LABEL FOR="cboMBOM1">��ü</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBOM" tag="1X" ID="cboMBOM2" VALUE="Y"><LABEL FOR="cboMBOM2">��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="cboMBOM" tag="1X" ID="cboMBOM3" VALUE="N"><LABEL FOR="cboMBOM3">�ƴϿ�</LABEL></TD>
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
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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

