
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

    lgIntFlgMode = parent.OPMD_CMODE					'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed

End Sub

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 

Sub SetDefaultVal()
	//txtQueryFrom.Text	= UNIGetFirstDay("<%=GetSvrDate%>", Parent.gDateFormat)
	txtQueryFrom.Text		= UniConvDateAToB("<%=dateadd("m","-1",GetSvrDate)%>" ,parent.gServerDateFormat,gDateFormat)
	txtQueryTo.Text		= UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,gDateFormat)
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
		Case 0, 1
			arrParam(0) = "품목팝업"
			arrParam(1) = "B_ITEM a, B_ITEM_BY_PLANT b"
			arrParam(2) = strCode
			arrParam(3) = ""
			arrParam(4) = " a.ITEM_CD = b.ITEM_CD "
			If Trim(cboItemAcct.value) <> "" Then
				arrParam(4) = arrParam(4) & " AND b.ITEM_ACCT = " & FilterVar(cboItemAcct.value, "''", "S")
			End If
			arrParam(5) = "품목"			
	
			arrField(0) = "a.ITEM_CD"
			arrField(1) = "a.ITEM_NM"
			arrField(2) = "a.SPEC"
			 
			arrHeader(0) = "품목코드"
			arrHeader(1) = "품목명"
			arrHeader(2) = "Spec"

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
				txtItemCdFrom.focus
			Case 1
				txtItemCdTo.focus

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
			txtItemCdFrom.focus
			txtItemCdFrom.value = arrRet(0)
			txtItemNmFrom.value = arrRet(1)
		Case 1
			txtItemCdTo.focus
			txtItemCdTo.value = arrRet(0)
			txtItemNmTo.value = arrRet(1)

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

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'======================================================================================================
'   Event Name : txtQueryFrom_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtQueryFrom_DblClick(Button)
    If Button = 1 Then
        fpDateTime1.Action = 7
    End If
End Sub

Sub txtQueryTo_DblClick(Button)
    If Button = 1 Then
        fpDateTime2.Action = 7
    End If
End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

	Err.clear
	
	Call CommonQueryRs("MINOR_CD, MINOR_NM", "B_MINOR", "MAJOR_CD = " & FilterVar("P1001", "''", "S") , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	Call SetCombo2(cboItemAcct ,lgF0  ,lgF1  ,Chr(11))
	
End Sub

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strDateFrom, strDateTo
	Dim strItemAcct, strItemCdFrom, strItemCdTo
	Dim strPlantCd
		
	SetPrintCond = False

	strDateFrom		= UNIConvDateToYYYYMMDD(txtQueryFrom.Text ,parent.gDateFormat, "")
	strDateTo		= UNIConvDateToYYYYMMDD(txtQueryTo.Text ,parent.gDateFormat, "")
	strItemAcct		= Trim(cboItemAcct.value)
	strItemCdFrom	= Trim(txtItemCdFrom.value)
	strItemCdTo		= Trim(txtItemCdTo.value)
	strPlantCd		= Trim(txtPlantCd.value)

	If strItemAcct = "" Then
		strItemAcct = "%"
	End If

	If strItemCdFrom = "" Then
		strItemCdFrom = " "
	End If

	If strItemCdTo = "" Then
		strItemCdTo = "ZZZZZZZZZZZZZZZZZZ"
	End If

	If strPlantCd = "" Then
		strPlantCd = "%"
	End If

	StrEbrFile	= "VI003MA1"
	
	StrUrl = StrUrl & "strDateFrom|"	& strDateFrom
	StrUrl = StrUrl & "|strDateTo|"		& strDateTo
	StrUrl = StrUrl & "|strItemAcct|"	& strItemAcct
	StrUrl = StrUrl & "|strItemCdFrom|"	& strItemCdFrom
	StrUrl = StrUrl & "|strItemCdTo|"	& strItemCdTo
	StrUrl = StrUrl & "|strPlantCd|"	& strPlantCd

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
	lgEBProcessbarOut = "T"
	EBActionA.menu.value = 0
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

'======================================================================================================
' Function Name :
' Function Desc :
'=======================================================================================================
Function DoSpExec(strPrintOpt)		'Preview, Print

	If ValidDateCheck(txtQueryFrom, txtQueryTo)	=	False	Then Exit	Function
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
' Function Name : MyBizASP1_OnReadyStateChange
' Function Desc : 
'========================================================================================

Sub MyBizASP1_onreadystatechange()
	If lgEBProcessbarOut = "T" Then		
	   Call LayerShowHide(0)
	   lgEBProcessbarOut = "F"  
	End  If   
End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'========================================================================================================
'   Event Name : txtItemCdFrom_onChange
'   Event Desc : 
'========================================================================================================


'========================================================================================================
'   Event Name : txtItemCdTo_onChange
'   Event Desc : 
'========================================================================================================
Sub txtItemCdTo_onChange()
	Dim IntRetCD
	Dim arrVal

	If txtItemCdTo.value = "" Then 
		txtItemNmTo.value	= ""
		Exit Sub
	End If

End Sub

Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"					' 팝업 명칭 
	arrParam(1) = "B_PLANT"					' TABLE 명칭 
	arrParam(2) = Trim(txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""						' Name Cindition
	arrParam(4) = ""						' Where Condition
	arrParam(5) = "공장"						' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"				' Field명(0)
    arrField(1) = "PLANT_NM"				' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlantCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	txtPlantCd.focus
	
End Function

Function SetPlantCd(ByVal arrRet)
	txtPlantCd.value = arrRet(0)
	txtPlantNm.value = arrRet(1)  
End Function


'========================================================================================
' Function Name : txtPlantCd_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtPlantCd_OnChange()
    Dim IntRetCd

    If txtPlantCd.value = "" Then
        txtPlantnm.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" PLANT_NM "," B_PLANT "," PLANT_CD="&filterVar(txtPlantCd.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 call DisplayMsgBox("971012","X", "공장","X")
			  txtPlantnm.value=""
			 txtPlantCd.focus
			
        Else
            txtPlantnm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
    
End Function

'========================================================================================
' Function Name : txtItemCdFrom_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtItemCdFrom_OnChange()
    Dim IntRetCd
    If txtItemCdFrom.value = "" Then
        txtItemNmFrom.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" ITEM_NM "," B_ITEM a, B_ITEM_BY_PLANT b "," a.ITEM_CD = b.ITEM_CD And a.ITEM_CD="&filterVar(txtItemCdFrom.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 call DisplayMsgBox("971012","X", txtItemCdFrom.ALT,"X")
			  txtItemNmFrom.value=""
			 txtItemCdFrom.focus
			
        Else
            txtItemNmFrom.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
    
End Function


'========================================================================================
' Function Name : txtItemCdTo_cd_OnChange
' Function Desc : 
'========================================================================================
Function txtItemCdTo_OnChange()
    Dim IntRetCd

    If txtItemCdTo.value = "" Then
        txtItemNmTo.value = ""
    ELSE    
        IntRetCd =  CommonQueryRs(" ITEM_NM "," B_ITEM a, B_ITEM_BY_PLANT b "," a.ITEM_CD = b.ITEM_CD And a.ITEM_CD="&filterVar(txtItemCdTo.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
			 call DisplayMsgBox("971012","X", txtItemCdTo.ALT,"X")
			 txtItemNmTo.value=""
			 txtItemCdTo.focus
			
        Else
            txtItemNmTo.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
    
End Function



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
			<TD CLASS="TD5E" NOWRAP>조회일자</TD>
			<TD CLASS="TD6E" NOWRAP><script language =javascript src='./js/vi003ma1_fpDateTime1_txtQueryFrom.js'></script>&nbsp;~&nbsp;
									<script language =javascript src='./js/vi003ma1_fpDateTime2_txtQueryTo.js'></script>
			</TD>
			<TD CLASS="TD5E" NOWRAP>품목계정</TD>
			<TD CLASS="TD6E" NOWRAP><SELECT NAME="cboItemAcct" tag="11X" STYLE="WIDTH:240px:" ALT="품목계정"><OPTION VALUE="" selected></OPTION></SELECT></TD>
		</TR>
		<TR>
			<TD CLASS="TD5E" NOWRAP>품목</TD>
			<TD CLASS="TD6E" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCdFrom" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="시작품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemFromCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup txtItemCdFrom.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmFrom" SIZE=20 tag="14">&nbsp;~&nbsp;
										<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCdTo" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="종료품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemToCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup txtItemCdTo.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNmTo" SIZE=20 tag="14">
			</TD>
			<TD CLASS="TD5E" NOWRAP>공장</TD>
			<TD CLASS="TD6E" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
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

