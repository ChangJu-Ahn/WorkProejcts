
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
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. 선 언 부 
'##########################################################################################################

'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!-- #Include file="../../inc/incEISComm.asp"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/button.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">				</SCRIPT>
<Script Language="JavaScript"	SRC="../../inc/incImage.js">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit						'☜: indicates that All variables must be declared in advance

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop
Dim intRetCD

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

<!-- #Include file="../../inc/lgvariables.inc" --> 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
   
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "V", "NOCOOKIE", "OA") %>
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "사원 Popup"
			arrParam(1) = "HAA010T"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "사번"			
	
			arrField(0) = "EMP_NO"
			arrField(1) = "NAME"
			arrField(2) = "RES_NO"
			 
			arrHeader(0) = "사번"
			arrHeader(1) = "성명"
			arrHeader(2) = "주민번호"

		Case Else
			Exit Function

	End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=540px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0
				txtEmpNo.focus

			Case Else
		End Select

		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 
Function SetPopup(Byval arrRet, Byval iWhere)
	
	Select Case iWhere
		Case 0
			txtEmpNo.focus
			txtEmpNo.value = arrRet(0)
			txtEmpNm.value = arrRet(1)

		Case Else
			Exit Function

	End Select

End Function


'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                           '⊙: Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")        '⊙: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")         '⊙: 조건에 맞는 Field locking
 
    Call InitVariables                            '⊙: Initializes local global Variables
    Call SetDefaultVal
                  
	Call InitComboBox
    
    Call SetToolbar("10000000000011")				'⊙: 버튼 툴바 제어 

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	
	Err.clear
	
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strEmpNo
		
	SetPrintCond = False

	strEmpNo	= Trim(txtEmpNo.value)

	StrEbrFile	= "VH003MA1"
	
	StrUrl = StrUrl & "strEmpNo|"	& strEmpNo

	SetPrintCond = True
	
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	'On Error Resume Next                                                    '☜: Protect system from crashing
    
    Dim StrUrl, StrEbrFile, ObjName

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Call LayerShowHide(0)
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	
	EBActionA.menu.value = 0
    lgEBProcessbarOut = "T"
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"ebr")	
			
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	'On Error Resume Next                                                    '☜: Protect system from crashing

    Dim StrUrl, StrEbrFile, ObjName
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Call LayerShowHide(0)
       Exit Function
    End If
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
		
End Function

'========================================================================================
' Function Name : DoSpExec
' Function Desc : 
'========================================================================================

Function DoSpExec(strPrintOpt)

    lgEBProcessbarOut = "F"

	' 화면 초기화 
	MyBizASP1.location.href = "../../blank.htm"
    
	If strPrintOpt = "Preview" Then
		Call LayerShowHide(1)
		Call FncBtnPreview() 
	ElseIf strPrintOpt = "Print" Then
		Call FncBtnPrint() 
	End If

End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================================
'   Event Name : txtEmpNo_onChange
'   Event Desc : 
'========================================================================================================
Sub txtEmpNo_onChange()
	Dim IntRetCD
	Dim arrVal

	If txtEmpNo.value = "" Then 
		txtEmpNm.value	= ""
		Exit Sub
	End If

	If CommonQueryRs("NAME", "HAA010T", " EMP_NO=  " & FilterVar(txtEmpNo.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		txtEmpNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("800006","X","X","X")  	'인사마스타에 등록된 사원이 아닙니다.
		txtEmpNo.focus
	End If

End Sub

'========================================================================================================
Sub MyBizASP1_onreadystatechange()
	If lgEBProcessbarOut = "T" Then		
	   Call LayerShowHide(0)
	   lgEBProcessbarOut = "F"  
	End  If   
End Sub


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
'       					6. Tag부 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">

    <%
        Call PrintTitle(Request("strASPMnuMnuNm"))
    %>

	<CENTER>
	<TABLE  <%=TABSTYLE01%> >
		<TR>
			<TD CLASS="TD5E" NOWRAP>사번</TD>
			<TD CLASS="TD6E" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtEmpNo" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="사번"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup txtEmpNo.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtEmpNm" SIZE=20 tag="14">
			</TD>
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

