
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : EIS
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Shin Hyun Ho
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!-- #Include file="../../inc/incEISComm.asp"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/button.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">				</SCRIPT>
<Script Language="JavaScript"	SRC="../../inc/incImage.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit						'��: indicates that All variables must be declared in advance

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop
Dim intRetCD

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID = "VC001MB1.asp"  


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    '---- Coding part--------------------------------------------------------------------    
End Sub

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	txtQueryDate.Text	= UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,gDateFormat)
	Call ggoOper.FormatDate(txtQueryDate, gDateFormat, 2)
	cboUnit.value = "1000"
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "V", "NOCOOKIE", "OA") %>
End Sub

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                           '��: Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
 
    Call InitVariables                            '��: Initializes local global Variables
                  
	Call InitComboBox
    
    Call SetDefaultVal
    
    Call SetToolbar("10000000000011")				'��: ��ư ���� ���� 

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'======================================================================================================
'   Event Name : txtQueryDate_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtQueryDate_DblClick(Button)
    If Button = 1 Then
        fpDateTime1.Action = 7
    End If
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

	Err.clear
	
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD = " & FilterVar("V0001", "''", "S") , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	Call SetCombo2(cboUnit ,lgF0  ,lgF1  ,Chr(11))
	
	'Call CommonQueryRs("BIZ_AREA_CD, BIZ_AREA_NM", "B_BIZ_AREA", "" , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	'Call SetCombo2(cboBizArea ,lgF0  ,lgF1  ,Chr(11))
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strSpId, strUnit,strYyyymm
		
	SetPrintCond = False

	strYyyymm	= txtQueryDate.year & right("0" & txtQueryDate.month,2)
	strSpId		= FilterVar(txtSpId.value,"","SNM")
	strUnit	  	= FilterVar(cboUnit.value,"","SNM")
	
	StrEbrFile	= "VC001MA1"

	StrUrl = StrUrl & "strYYYYMM|" & strYyyymm
	StrUrl = StrUrl & "|strSpId|" & strSpId
	StrUrl = StrUrl & "|strUnit|" & strUnit

	SetPrintCond = True
	
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '��: Protect system from crashing
    
    Dim StrUrl, StrEbrFile, ObjName

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	
	EBActionA.menu.value = 0
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"EBR")	
			
	Call LayerShowHide(0)
			
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	'On Error Resume Next                                                    '��: Protect system from crashing

    Dim StrUrl, StrEbrFile, ObjName
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

'======================================================================================================
' Function Name :
' Function Desc :
'=======================================================================================================
Function DoSpExec(strPrintOpt)		'Preview, Print
	Dim strVal       
	Dim strYyyyMm, strBizAreaCd

	If Not chkField(Document, "1") Then                             '��: Check contents area
		Exit Function
	End If

	' ȭ�� �ʱ�ȭ 
	MyBizASP1.location.href = "../../blank.htm"

	strYyyyMm	= txtQueryDate.year & right("0" & txtQueryDate.month,2)
	strBizAreaCd	= txtBizAreaCd.value
	
	if trim(strBizAreaCd) = "" then
		strBizAreaCd	= "%"
	end if
	
	' �̸������ �۾��� �޼��� ó�� 
	If strPrintOpt = "Preview" Then
		Call LayerShowHide(1) 
	End If
	    
	strVal = BIZ_PGM_ID & "?txtMode="	& Parent.UID_M0002
	strVal = strVal & "&strPrintOpt="	& strPrintOpt
	strVal = strVal & "&strYyyyMm="		& strYyyyMm
	strVal = strVal & "&strBizAreaCd="	& strBizAreaCd

	Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

Function OpenBizAreaPopUp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������˾�"                             ' Popup Name
	arrParam(1) = "B_BIZ_AREA"                                        ' Table Name
	arrParam(2) = txtBizAreaCd.value                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = ""                                              ' Where Condition
	arrParam(5) = "������ڵ�"
	
    arrField(0) = "BIZ_AREA_CD"                                      ' Field��(0)
    arrField(1) = "BIZ_AREA_NM"                                      ' Field��(1)
    
    arrHeader(0) = "������ڵ�"	                              ' Header��(0)
    arrHeader(1) = "������"                               ' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		txtBizAreaCd.focus
		Exit Function
	Else
		Call SetBizArea(arrRet)
	End If	


End Function

Sub SetBizArea(Byval arrRet)
	  txtBizAreaCd.value = arrRet(0)
	  txtBizAreaNm.value = arrRet(1)
End Sub

'========================================================================================================
'========================================================================================================
'   Event Name : txtBizAreaCd_onChange
'   Event Desc : 
'========================================================================================================
Sub txtBizAreaCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If txtBizAreaCd.value = "" Then
	   txtBizAreaNm.value = ""
	   txtBizAreaCd.focus 
	Exit Sub
	End if 

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(txtBizAreaCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		txtBizAreaNm.value= Trim(arrVal(0)) 
	Else
		txtBizAreaCd.value = "":		txtBizAreaNm.value = ""
		IntRetCD = DisplayMsgBox("124200","X","X","X")  	'%1 ����� �����Ͱ� �������� �ʽ��ϴ�.
		txtBizAreaCd.focus
	End If
End Sub


'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

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
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">

    <%
        Call PrintTitle(Request("strASPMnuMnuNm"))
    %>

	<CENTER>
	<TABLE  <%=TABSTYLE01%> >
		<TR>
			<TD CLASS="TD5E" NOWRAP>��ȸ����</TD>
			<TD CLASS="TD6E" NOWRAP><script language =javascript src='./js/vc001ma1_fpDateTime1_txtQueryDate.js'></script>&nbsp;</TD>
			<TD CLASS="TD5E" NOWRAP>�ݾ״���</TD>
			<TD CLASS="TD6E" NOWRAP><SELECT NAME="cboUnit" tag="12X" STYLE="WIDTH:120px:" ALT="�ݾ״���"><OPTION VALUE="" selected></OPTION></SELECT></TD>
		</TR>
		<TR>
			<TD CLASS="TD5E" NOWRAP>�����</TD>
			<TD CLASS="TD6E" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�"><IMG SRC="../../image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaPopUp()"> <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=18 tag="14X" ALT="������"></TD>
			<TD CLASS="TD5E" NOWRAP></TD>
			<TD CLASS="TD6E" NOWRAP></TD>
		</TR>
	</TABLE>
	
	<TABLE width=1016  height=476 cellspacing=0 cellpadding=0 border=0>
		<TR>
			<TD><IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=auto framespacing=0 marginwidth=0 marginheight=0 ></IFRAME></TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=1><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
			</TD>
		</TR>
	</table>
	</center>

	<INPUT TYPE="HIDDEN" NAME="txtSpId" tag="24" TABINDEX = "-1">

	<FORM NAME="EBActionA" ID="EBAction" TARGET="MyBizASP1" METHOD="POST"  scroll=yes> 
		<input type="hidden" name="menu" value=0 > 
		<input type="hidden" name="id" > 
		<input type="hidden" name="pw" >
		<input type="hidden" name="doc" > 
		<input type="hidden" name="form" > 
		<input type="hidden" name="runvar" > 
	</FORM>

	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
	
	<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
		<INPUT TYPE="HIDDEN" NAME="uname"		TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="dbname"		TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="filename"	TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="condvar"		TABINDEX = "-1" >
		<INPUT TYPE="HIDDEN" NAME="date"		TABINDEX = "-1" >	
	</FORM>
</BODY>
</HTML>
