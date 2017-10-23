<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 매출채권순위출력 
'*  3. Program ID           : s5111oa3
'*  4. Program Name         : 매출채권순위출력 
'*  5. Program Desc         : 매출채권순위 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/18
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : 손범열 
'* 10. Modifier (Last)      : Hwang Seong Bae
'* 11. Comment              : 표준반영 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Dim IsOpenPop          

'=========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.BillFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.BillToDt.Text = EndDate
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "OA") %>
End Sub

'=========================================
Sub Form_Load()

  	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
End Sub

'==========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'==========================================
Sub BillFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.BillFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.BillFromDt.Focus
    End If
End Sub

'==========================================
Sub BillToDt_DblClick(Button)
    If Button = 1 Then
        frm1.BillToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.BillToDt.Focus
    End If
End Sub

'========================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================
 Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================
Function btnPreview_OnClick()
	Call BtnPrint("N")
End Function

'========================================
Function btnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'========================================
Function BtnPrint(ByVal pvStrPrint) 
	Dim iStrUrl
    
    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.BillFromDt, frm1.BillToDt) = False Then Exit Function

	iStrUrl =			"BillFromDt|" & UniConvDateToYYYYMMDD(frm1.BillFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	iStrUrl = iStrUrl & "|BillToDt|" & UniConvDateToYYYYMMDD(frm1.BillToDt.Text,Parent.gDateFormat,Parent.gServerDateType)
     	
	'SON 매출채권순위를 주문처별, 영업조직별로 볼수 있게 고침	
	If frm1.Rb_WK1.checked = True Then
		ObjName = AskEBDocumentName("s5111oa3","ebr")

	ELSEIF frm1.Rb_WK2.checked = True Then
     	ObjName = AskEBDocumentName("s5111oa7","ebr")
    END IF
    
	If pvStrPrint = "N" Then
		' 미리보기 
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		' 출력 
		Call FncEBRprint(EBAction, ObjName, iStrUrl)
	End If
    
End Function

'========================================
Function FncExit()
	FncExit = True
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권순위</font></td>
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
									<TD CLASS="TD5" NOWRAP>매출채권일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s5111oa3_fpDateTime1_BillFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s5111oa3_fpDateTime2_BillToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>순위</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK1 Checked><LABEL FOR=Rb_WK1>주문처별</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2><LABEL FOR=Rb_WK2>영업조직별</LABEL></TD>
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
						<TD>
						    <BUTTON NAME="btnPreview" CLASS="CLSSBTN" Flag=1>미리보기</BUTTON>&nbsp;
						    <BUTTON NAME="btnPrint" CLASS="CLSSBTN" Flag=1>인쇄</BUTTON>
						</TD>
					</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX ="-1">
    <input type="hidden" name="dbname" TABINDEX ="-1">
    <input type="hidden" name="filename" TABINDEX ="-1">
    <input type="hidden" name="condvar" TABINDEX ="-1">
	<input type="hidden" name="date" TABINDEX ="-1">
</FORM>
</BODY>
</HTML>
