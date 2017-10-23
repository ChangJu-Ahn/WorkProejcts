<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p4316oa1.asp
'*  4. Program Name         : ��������û�����(������)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/01/10
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<% '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################%>
<% '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* %>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<%'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>

<%'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================%>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--��:Print Program needs this vbs file-->
<SCRIPT LANGUAGE = VBSCRIPT>

Option Explicit                                                             '��: indicates that All variables must be declared in advance

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

Dim LocSvrDate
Dim StartDate
Dim EndDate

LocSvrDate = "<%=GetSvrDate%>"
StartDate = UniConvDateAToB(LocSvrDate,parent.gServerDateFormat,parent.gDateFormat)     	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
EndDate = parent.UNIDateAdd("M",1,StartDate,parent.gDateFormat)							    '��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

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
'========================================================================================================= %>
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
	frm1.txtPlanStartDt1.Text	= StartDate
	frm1.txtPlanStartDt2.Text	= EndDate
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt1_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPlanStartDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanStartDt1.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanStartDt1.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt2_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtPlanStartDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtPlanStartDt2.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtPlanStartDt2.Focus
    End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
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
    
	Call ggoOper.FormatField(Document, "x",parent.ggStrIntegeralPart, parent.ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)        
	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
    
    Call SetToolbar("10000000000011")												'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtProdtOrdNo1.focus
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

	Dim strEbrFile
    Dim objName
	
    Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If parent.ValidDateCheck(frm1.txtPlanStartDt1, frm1.txtPlanStartDt2) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))	
	
	if frm1.txtProdtOrdNo1.value = "" then 
		var2 = "0"	
	else
		var2 = UCase(Trim(frm1.txtProdtOrdNo1.value))
	End If
	
	if frm1.txtProdtOrdNo2.value = "" then 
		var3 = "zzzzzzzzzzzzzzzzzz"	
	else
		var3 = UCase(Trim(frm1.txtProdtOrdNo2.value))
	End If
	
	var4 = UniConvDateAToB(frm1.txtPlanStartDt1.Text,parent.gDateFormat,parent.gServerDateFormat) 
	var5 = UniConvDateAToB(frm1.txtPlanStartDt2.Text,parent.gDateFormat,parent.gServerDateFormat)
	
	strUrl = strUrl & "plant_cd|" & var1
	strUrl = strUrl & "|fr_order_no|" & var2
	strUrl = strUrl & "|to_order_no|" & var3
	strUrl = strUrl & "|fr_date|" & var4
	strUrl = strUrl & "|to_date|" & var5
	
	strEbrFile = "p4316oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr") 
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
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
	
	Dim strEbrFile
    Dim objName
	
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	
	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then							'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If parent.ValidDateCheck(frm1.txtPlanStartDt1, frm1.txtPlanStartDt2) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	If frm1.txtProdtOrdNo1.value = "" then 
		var2 = "0"	
	else
		var2 = UCase(Trim(frm1.txtProdtOrdNo1.value))
	End If
	
	if frm1.txtProdtOrdNo2.value = "" then 
		var3 = "zzzzzzzzzzzzzzzzzz"	
	else
		var3 = UCase(Trim(frm1.txtProdtOrdNo2.value))
	End If
	
	var4 = UniConvDateAToB(frm1.txtPlanStartDt1.Text,parent.gDateFormat,parent.gServerDateFormat) 
	var5 = UniConvDateAToB(frm1.txtPlanStartDt2.Text,parent.gDateFormat,parent.gServerDateFormat)
	
	strUrl = strUrl & "plant_cd|" & var1
	strUrl = strUrl & "|fr_order_no|" & var2
	strUrl = strUrl & "|to_order_no|" & var3
	strUrl = strUrl & "|fr_date|" & var4
	strUrl = strUrl & "|to_date|" & var5 
	
	strEbrFile = "p4316oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================
Function FncQuery()
	On Error Resume Next                                                    '��: Protect system from crashing	
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)									<%'��:ȭ�� ����, Tab ���� %>
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()                                         <%'��: Protect system from crashing%>
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
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

	If IsOpenPop = True Then
	 Exit Function
	End If 

	IsOpenPop = True

	arrParam(0) = "�����˾�"			' �˾� ��Ī 
	arrParam(1) = "B_PLANT"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""						' Name Cindition
	arrParam(4) = ""						' Where Condition
	arrParam(5) = "����"				' TextBox ��Ī 
	
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

'------------------------------------------  OpenFromProdOrderNo()  -------------------------------------------------
'	Name : OpenFromProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenFromProdOrderNo()
	
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdtOrdNo1.className) = UCase(parent.UCN_PROTECTED) Then
	 Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call parent.DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtPlanStartDt1.Text
	arrParam(2) = frm1.txtPlanStartDt2.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrdNo1.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetFromProdtOrdNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrdNo1.focus
	
End Function

'------------------------------------------  OpenToProdOrderNo()  -------------------------------------------------
'	Name : OpenToProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenToProdOrderNo()
	
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdtOrdNo2.className) = UCase(parent.UCN_PROTECTED) Then
	 Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call parent.DisplayMsgBox("971012","X" , "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtPlanStartDt1.Text
	arrParam(2) = frm1.txtPlanStartDt2.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdtOrdNo2.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent ,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetToProdtOrdNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtProdtOrdNo2.focus
	
End Function

'------------------------------------------  SetPlantCd()  --------------------------------------------------
'	Name : SetPlantCd()
'	Description : Plant  Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)  
End Function

'------------------------------------------  SetProdtOrdNo()  --------------------------------------------------
'	Name : SetProdtOrdNo()
'	Description : SetProdtOrdNo1 Popup���� return�� �� 
'---------------------------------------------------------------------------------------------------------
Function SetFromProdtOrdNo(ByVal arrRet)
	frm1.txtProdtOrdNo1.value = arrRet(0) 
End Function
		
Function SetToProdtOrdNo(ByVal arrRet)	
	frm1.txtProdtOrdNo2.value = arrRet(0) 
End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����û��(������)</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>		
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="x2xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="x4" ALT="�����">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�۾���ȹ����</TD>
									<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p4316oa1_I722931295_txtPlanStartDt1.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/p4316oa1_I230755758_txtPlanStartDt2.js'></script>	
								</TR>						
								<TR>
									<TD CLASS="TD5" NOWRAP>����������ȣ</TD>
									<TD CLASS="TD656" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtProdtOrdNo1" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="�۾����ù�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdtOrdNo1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromProdOrderNo()">&nbsp;~&nbsp;
									</TD>
								</TR>						
								<TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD656" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtProdtOrdNo2" SIZE=18 MAXLENGTH=18 tag="x1xxxU" ALT="�۾����ù�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdtOrdNo2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToProdOrderNo()">
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
				  <TD WIDTH = 10></TD>
		          <TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON></TD>		
	            </TR>
	        </TABLE>
	    </TD>    
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
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
