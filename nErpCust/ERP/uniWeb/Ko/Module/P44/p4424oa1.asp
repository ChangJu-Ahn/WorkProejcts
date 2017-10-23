<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : production
'*  2. Function Name        : 
'*  3. Program ID           : p4424oa1
'*  4. Program Name         : (p)���ְ����񳻿� ��� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001.11.22
'*  9. Modifier (First)     : Jaehyun Chen
'* 10. Modifier (Last)      : Jaehyun Chen
'* 11. Comment              :
'*                          : Order Number���� �ڸ��� ����(2003.04.14) Park Kye Jin
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" --> 

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>              <!--��:Print Program needs this vbs file-->
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance

'=========================================================================================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop

Dim StartDate
Dim EndDate
Dim strYear, strMonth, strDay

Call ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)

EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)	    	'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")	   '��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
End Sub

'=========================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtStartDt.Text = StartDate
	frm1.txtEndDt.Text = EndDate
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "P","NOCOOKIE","OA") %>		
End Sub

'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "X",parent.ggStrIntegeralPart, parent.ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
	Call SetDefaultVal
    Call InitVariables														'��: Initializes local global variables
    Call SetToolbar("10000000000011")										'��: ��ư ���� ���� 
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	End If		
	
	frm1.txtfrBpCd.focus 
	Set gActiveElement = document.activeElement
	
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
	Dim var6
	Dim var7
	Dim var8
	
	dim strUrl
	dim arrParam, arrField, arrHeader

'--------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� ���� 
'--------------------------------------------------------------
	Dim lngPos
	Dim intCnt
'--- End : Declare Variables ----------------------------------
			
	If frm1.txtFrBpCd.value = "" Then
		frm1.txtFrBpCd.value = "" 
	End If	
	
	If frm1.txtToBpCd.value = "" Then
		frm1.txtToBpNm.value = "" 
	End If	
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtFrWcCd.value = "" Then
		frm1.txtFrWcNm.value = "" 
	End If
	
	If frm1.txtToWcCd.value = "" Then
		frm1.txtToWcNm.value = "" 
	End If
	
    Call BtnDisabled(1)	
	
	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If parent.ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF

	If frm1.txtFrBpCd.value = "" Then
		var1 = "0"
	Else
		var1 = Trim(frm1.txtFrBpCd.value)
	End If
	
	If frm1.txtToBpCd.value = "" Then
		var2 = "zzzzzzzzzz"
	Else
		var2 = Trim(frm1.txtToBpCd.value)
	End If
	
	var3 = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	var4 = UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If frm1.txtPlantCd.value = "" Then
		var5 = "0"
		var6 = "zzzz"		
	Else
		var5 = Trim(frm1.txtPlantCd.value)
		var6 = Trim(frm1.txtPlantCd.value)
	End If
	
	If frm1.txtFrWcCd.value = "" Then
		var7 = "0"		
	Else
		var7 = Trim(frm1.txtFrWcCd.value)
	End If
	
	If frm1.txtToWcCd.value = "" Then
		var8 = "zzzzzzz"	
	Else
		var8 = Trim(frm1.txtToWcCd.value)
	End If

	strUrl = strUrl & "fr_bp_cd|" & var1 
	strUrl = strUrl & "|to_bp_cd|" & var2
	strUrl = strUrl & "|fr_start_dt|" & var3 
	strUrl = strUrl & "|to_start_dt|" & var4
	strUrl = strUrl & "|fr_plant_cd|" & var5 
	strUrl = strUrl & "|to_plant_cd|" & var6 
	strUrl = strUrl & "|fr_wc_cd|" & var7 
	strUrl = strUrl & "|to_wc_cd|" & var8 
	
	strEbrFile = "p4424oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	Call FncEBRprint(EBAction, objName, strUrl)
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
	Dim var6
	Dim var7
	Dim var8	
	
	dim strUrl
	dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "x") Then									'��: This function check indispensable field
		Call BtnDisabled(0)	
       Exit Function
    End If
	
	If frm1.txtToBpCd.value = "" Then
		frm1.txtToBpNm.value = "" 
	End If	
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtFrWcCd.value = "" Then
		frm1.txtFrWcNm.value = "" 
	End If
	
	If frm1.txtToWcCd.value = "" Then
		frm1.txtToWcNm.value = "" 
	End If
	
	If parent.ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	

	If frm1.txtFrBpCd.value = "" Then
		var1 = "0"
	Else
		var1 = Trim(frm1.txtFrBpCd.value)
	End If
	
	If frm1.txtToBpCd.value = "" Then
		var2 = "zzzzzzzzzz"
	Else
		var2 = Trim(frm1.txtToBpCd.value)
	End If
	
	var3 = UniConvDateAToB(frm1.txtStartDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	var4 = UniConvDateAToB(frm1.txtEndDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If frm1.txtPlantCd.value = "" Then
		var5 = "0"
		var6 = "zzzz"		
	Else
		var5 = Trim(frm1.txtPlantCd.value)
		var6 = Trim(frm1.txtPlantCd.value)
	End If
	
	If frm1.txtFrWcCd.value = "" Then
		var7 = "0"		
	Else
		var7 = Trim(frm1.txtFrWcCd.value)
	End If
	
	If frm1.txtToWcCd.value = "" Then
		var8 = "zzzzzzz"	
	Else
		var8 = Trim(frm1.txtToWcCd.value)
	End If

	strUrl = strUrl & "fr_bp_cd|" & var1 
	strUrl = strUrl & "|to_bp_cd|" & var2
	strUrl = strUrl & "|fr_start_dt|" & var3 
	strUrl = strUrl & "|to_start_dt|" & var4
	strUrl = strUrl & "|fr_plant_cd|" & var5 
	strUrl = strUrl & "|to_plant_cd|" & var6 
	strUrl = strUrl & "|fr_wc_cd|" & var7 
	strUrl = strUrl & "|to_wc_cd|" & var8 
	
	strEbrFile = "p4424oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	Call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = parent.DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'------------------------------------------  OpenBizPartner()  -------------------------------------------------
'	Name : OpenBizparener()
'	Description : BpPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromBizPartner()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = frm1.txtFrBpCd.value 
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    arrField(2) = "BP_TYPE"
    arrField(3) = ""	
        
    arrHeader(0) = "BP"		
    arrHeader(1) = "BP��"		
    arrHeader(2) = "Bp ����"		
    arrHeader(3) = ""
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtFrBpCd.Value    = arrRet(0)		
		frm1.txtFrBpNm.Value    = arrRet(1)	
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtFrBpCd.focus
	
End Function

'------------------------------------------  OpenBizPartner()  -------------------------------------------------
'	Name : OpenBizparener()
'	Description : BpPopup
'--------------------------------------------------------------------------------------------------------- 
Function OpenToBizPartner()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó�˾�"	
	arrParam(1) = "B_BIZ_PARTNER"				
	arrParam(2) = frm1.txtToBpCd.value 
	arrParam(3) = ""
	arrParam(4) = "" 
	arrParam(5) = "����ó"			
	
    arrField(0) = "BP_CD"	
    arrField(1) = "BP_NM"	
    arrField(2) = "BP_TYPE"
    arrField(3) = ""	
        
    arrHeader(0) = "BP"		
    arrHeader(1) = "BP��"		
    arrHeader(2) = "Bp ����"		
    arrHeader(3) = ""
        
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtToBpCd.Value    = arrRet(0)		
		frm1.txtToBpNm.Value    = arrRet(1)	
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtToBpCd.focus
	
End Function

'------------------------------------------  OpenPlantCd()  ----------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"				' �˾� ��Ī 
	arrParam(1) = "B_PLANT"						' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "����"					' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"					' Field��(0)
    arrField(1) = "PLANT_NM"					' Field��(1)
    
    arrHeader(0) = "����"					' Header��(0)
    arrHeader(1) = "�����"					' Header��(1)
        
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)  
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromConWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If frm1.txtPlantCd.value= "" Then
		Call parent.DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtFrWCCd.Value)									' Code Condition
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " and INSIDE_FLG = " & FilterVar("N", "''", "S") & " " ' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtFrWCCd.Value    = arrRet(0)		
		frm1.txtFrWCNm.Value    = arrRet(1)		
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtFrWCCd.focus
		
End Function

'------------------------------------------  OpenConWC()  -------------------------------------------------
'	Name : OpenConWC()
'	Description : Condition Work Center PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenToConWC()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If frm1.txtPlantCd.value= "" Then
		Call parent.DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�۾����˾�"											' �˾� ��Ī 
	arrParam(1) = "P_WORK_CENTER"											' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtToWCCd.Value)									' Code Condition
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " and INSIDE_FLG = " & FilterVar("N", "''", "S") & " " ' Where Condition
	arrParam(5) = "�۾���"												' TextBox ��Ī 
	
    arrField(0) = "WC_CD"													' Field��(0)
    arrField(1) = "WC_NM"													' Field��(1)
    
    arrHeader(0) = "�۾���"												' Header��(0)
    arrHeader(1) = "�۾����"											' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtToWCCd.Value    = arrRet(0)		
		frm1.txtToWCNm.Value    = arrRet(1)		
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtToWCCd.focus
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���ְ����񳻿����</font></td>
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
								    <TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFrBpCd" SIZE=10 MAXLENGTH=10 tag="X1XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBPCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromBizPartner()"> <INPUT TYPE=TEXT ID="txtFrBpNm" NAME="arrCond" tag="X4">&nbsp;~</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtToBpCd" SIZE=10 MAXLENGTH=10 tag="X1XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBPCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToBizPartner()"> <INPUT TYPE=TEXT ID="txtToBpNm" NAME="arrCond" tag="X4"></TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="X1XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="�����">
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>�����۾���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFrWCCd" SIZE=7 MAXLENGTH=7 tag="X1XXXU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromConWC()"> <INPUT TYPE=TEXT ID="txtFrWCNm" NAME="arrCond" tag="X4">&nbsp;~</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtToWCCd" SIZE=7 MAXLENGTH=7 tag="X1XXXU" ALT="�۾���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWCCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToConWC()"> <INPUT TYPE=TEXT ID="txtToWCNm" NAME="arrCond" tag="X4"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�԰���</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p4424oa1_I247757798_txtStartDt.js'></script>								
										&nbsp;~&nbsp; 
										<script language =javascript src='./js/p4424oa1_I146974492_txtEndDt.js'></script>								
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
