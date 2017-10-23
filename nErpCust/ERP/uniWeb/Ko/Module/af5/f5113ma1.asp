<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1 %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Account Management
'*  3. Program ID           : f5113ma1.asp
'*  4. Program Name         : ������ǥ���Һ���� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/14
'*  8. Modified date(Last)  : 2001/01/03
'*  9. Modifier (First)     : Hersheys
'* 10. Modifier (Last)      : Hersheys
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            
'********************************************************************************************** -->
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
<!-- #Include file="../../inc/IncServer.asp"  -->				<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->

<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--
'=============================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '��: indicates that All variables must be declared in advance 

'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################


'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* 


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

Const BIZ_PGM_ID = "a6108mb1.asp"  

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											'��: �����Ͻ� ���� ASP���� �����ϹǷ� Public 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
'Dim cboOldVal          
Dim IsOpenPop          
'Dim lgCboKeyPress      
'Dim lgOldIndex								
'Dim lgOldIndex2        

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'*****************************************  2.1 Pop-Up �Լ�   ********************************************
'	���: Pop-Up 
'********************************************************************************************************* 

'===========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================== 

'++++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function FncBtnPrint() 
	On Error Resume Next
	
	Dim Var1
	Dim Var2

	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	
    lngPos = 0	

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
	var1 = Replace(frm1.fpDateTime1.text,gComDateType,gServerDateType)
	var2 = Replace(frm1.fpDateTime2.text,gComDateType,gServerDateType)
	
	For intCnt = 1 To 3
		lngPos = instr(lngPos + 1, GetUserPath, "/")
	Next

	StrEbrFile = "f5113ma1.ebr"

	strUrl = Left(GetUserPath, lngPos - 1)
	StrUrl = StrUrl & gEbEnginePath & "ExecuteWinReportForPrint?"
	StrUrl = StrUrl & "uname="    & gEbUserName   & "&"
	StrUrl = StrUrl & "dbname="   & gEbDbName     & "&"
	StrUrl = StrUrl & "filename=" & gEbPkgRptPath & "\" & gLang & "\ebr\" & StrEbrFile

	StrUrl = StrUrl & "&condvar=FromIssueDt|" & var1
	StrUrl = StrUrl & "|ToIssueDt|"	          & var2
	StrUrl = StrUrl & "|&date=-2"
		
	MyBizASP.location.href = strUrl
	
End Function

Function FncBtnPreview()
	On Error Resume Next
	
	Dim Var1
	Dim Var2

	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	
	var1 = Replace(frm1.fpDateTime1.text,gComDateType,gServerDateType)
	var2 = Replace(frm1.fpDateTime2.text,gComDateType,gServerDateType)
	
	StrEbrFile = "f5113ma1.ebr"

	StrUrl = gEbEnginePath & "ExecuteWinReport?"
	StrUrl = StrUrl & "uname="    & gEbUserName   & "&"
	StrUrl = StrUrl & "dbname="   & gEbDbName     & "&"
	StrUrl = StrUrl & "filename=" & gEbPkgRptPath & "\" & gLang & "\ebr\" & StrEbrFile

	StrUrl = StrUrl & "&condvar=FromIssueDt|" & var1
	StrUrl = StrUrl & "|ToIssueDt|"			  & var2
	StrUrl = StrUrl & "|&date=-2"
		
	Window.ShowModalDialog strUrl, Array(arrParam, arrField, arrHeader), _
		"dialogWidth=1200px; dialogHeight=800px; center: Yes; help: No; resizable: Yes; status: No;"
	
End Function


'##########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'#########################################################################################################

'*****************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'===========================================  3.1.1 Form_Load()  =========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    'Call InitVariables																'��: Initializes local global variables
    'Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,gComNum1000,gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000111")
    'Call SetDefaultVal

	frm1.txtFromIssueDt.focus 
	frm1.txtFromIssueDt.Value = Now
	frm1.txtToIssueDt.Value   = Now
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt.Action = 7
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt.Action = 7
    End If
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	

<%
'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------
%>

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
								<td background="../../image/table/seltab_up_bg.gif"><img src="../../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������ǥ���Һ����</font></td>
								<td background="../../image/table/seltab_up_bg.gif" align="right"><img src="../../image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">������</TD>
								<TD CLASS="TD6"><script language =javascript src='./js/f5113ma1_fpDateTime1_txtFromIssueDt.js'></script>
												 &nbsp;~&nbsp;
											    <script language =javascript src='./js/f5113ma1_fpDateTime2_txtToIssueDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPrint" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPrint()">���</BUTTON>&nbsp;<BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPreview()">�̸�����</BUTTON></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
