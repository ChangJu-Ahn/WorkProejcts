<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 거래명세표 
'*  3. Program ID           : S4111oa3
'*  4. Program Name         : 거래명세표 
'*  5. Program Desc         : 거래명세표 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/07/18
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : 손범열 
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim IsOpenPop          

'=========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.DN_NO.focus 
	frm1.cboApType.value = "공급받는자용"
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=========================================
Function OpenConPop()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("S4111PA3")

	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111PA3", "x")
		IsOpenPop = False
		exit Function
	end if
			
	IsOpenPop = True
	
	strRet = window.showModalDialog(iCalledAspName, array(window.parent), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

    frm1.DN_NO.focus 
		
	If strRet <> "" Then
		frm1.DN_NO.value = strRet		
	End If	

End Function

'=========================================
Sub InitComboBox()
	With frm1
		Call SetCombo(frm1.cboApType, "공급받는자용","공급받는자용")
		Call SetCombo(frm1.cboApType, "공급자용","공급자용")                              
	     .cboApType.value = "공급받는자용"
    End With
End Sub

'=========================================
Sub Form_Load()

    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
    Call InitComboBox
End Sub

'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
   
End Sub

'=========================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'=========================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================
Function BtnPreview_OnClick()
	Call BtnPrint("N")
End Function

'========================================
Function BtnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'=====================================================
Function BtnPrint(ByVal pvStrPrint) 

    If Not chkField(Document, "1") Then Exit Function

	Dim iStrUrl
    
    ' 출고유형 
	If Trim(frm1.DN_NO.value) = "" Then
		iStrUrl = "DN_NO|%"
	Else
		iStrUrl = "DN_NO|" & Replace(UCase(Trim(frm1.DN_NO.value)), "'" ,  "''")
	End If

	if frm1.cboAptype.Value = "공급받는자용" then
		OBjName = AskEBDocumentName("s4111oa4","ebr")    
	else
		OBjName = AskEBDocumentName("s4111oa3","ebr")    
	End If

	If pvStrPrint = "N" Then
		' 미리보기 
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		' 출력 
		Call FncEBRprint(EBAction, ObjName, iStrUrl)
	End If
		
End Function

'=====================================================
Function FncExit()
	FncExit = True
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>거래명세표출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>출하번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="DN_NO" ALT="출하번호" SIZE=20 MAXLENGTH=18 TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPop" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPop"><div style="display:none"><input type="text" name="none"></div></TD>
									                     
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboApType" ALT="거래명세표" STYLE="Width: 150px;" tag="22"></SELECT></TD>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
					<TD valign=top>
						<BUTTON NAME="BtnPreview" CLASS="CLSSBTN" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="BtnPrint" CLASS="CLSSBTN" Flag=1>인쇄</BUTTON>
					</TD>
				</TR> 	
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> SRC= "../../blank.htm" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
