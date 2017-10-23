
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : Template
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2004/05/
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho Ig sung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    
End Sub

Sub SetDefaultVal()


	Dim strFrDate, strToDate
'	strFrDate	= Parent.gFiscStart		' 당기시작일 
	strToDate	= "<%=GetSvrDate%>"		' 현재일 

	frm1.txtFrYyyymm.text	= UNIGetFirstDay(strToDate, Parent.gDateFormat)
	frm1.txtToYyyymm.text	= UniConvDateAToB(strToDate,Parent.gServerDateFormat,Parent.gDateFormat)

'	Call ggoOper.FormatDate(frm1.txtFrYyyymm, Parent.gDateFormat, 1)
'	Call ggoOper.FormatDate(frm1.txtToYyyymm, Parent.gDateFormat, 1)
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

' 모듈 코드로 입력 예:원가 "C"
<% Call loadInfTB19029A("Q", "B", "NOCOOKIE", "OA") %>
End Sub

Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitVariables
    Call SetDefaultVal

    Call SetToolbar("10000000000011")
    
    frm1.txtFrYyyymm.focus 
    Set gActiveElement = document.activeElement	
    
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

' 날짜에서 엔터키 입력시 미리보기 실행 
Sub txtFrYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrYyyymm.focus
	End If
End Sub

Sub txtToYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtToYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToYyyymm.focus
	End If
End Sub

Sub txtFrYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Sub txtToYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Function FncQuery()
	' 엔터키 입력시 미리보기 실행 
    FncBtnPreview()
End Function

Function SetPrintCond(StrEbrFile, strUrl)

	Dim strTrFrDt, strTrToDt, strItemAcct, strPlantCd, strFrItemCd, strToItemCd
	Dim	strFrYear, strFrMonth, strFrDay, strToYear, strToMonth, strToDay

	SetPrintCond = False

	if frm1.PrintOpt1.checked = True Then
		if frm1.ModuleOpt1.checked = True then
			StrEbrFile = "b7000oa1a"
		elseif frm1.ModuleOpt2.checked = True then
			StrEbrFile = "b7000oa1c"
		elseif frm1.ModuleOpt3.checked = True then
			StrEbrFile = "b7000oa1e"
		elseif frm1.ModuleOpt4.checked = True then
			StrEbrFile = "b7000oa1g"
		elseif frm1.ModuleOpt5.checked = True then
			StrEbrFile = "b7000oa1i"
		elseif frm1.ModuleOpt6.checked = True then
			StrEbrFile = "b7000oa1k"
		else
			Exit Function
		end if
	elseif frm1.PrintOpt2.checked = True Then
		if frm1.ModuleOpt1.checked = True then
			StrEbrFile = "b7000oa1b"
		elseif frm1.ModuleOpt2.checked = True then
			StrEbrFile = "b7000oa1d"
		elseif frm1.ModuleOpt3.checked = True then
			StrEbrFile = "b7000oa1f"
		elseif frm1.ModuleOpt4.checked = True then
			StrEbrFile = "b7000oa1h"
		elseif frm1.ModuleOpt5.checked = True then
			StrEbrFile = "b7000oa1j"
		elseif frm1.ModuleOpt6.checked = True then
			StrEbrFile = "b7000oa1l"
		else
			Exit Function
		end if
	else
		Exit Function
	end if

	' 시작월 종료월 비교 
    If CompareDateByFormat(frm1.txtFrYyyymm.Text,frm1.txtToYyyymm.Text,frm1.txtFrYyyymm.Alt,frm1.txtToYyyymm.Alt, _
	 "970024", frm1.txtFrYyyymm.UserDefinedFormat,Parent.gComDateType, true)=False then
		frm1.txtToYyyymm.Focus
		Exit Function
	End If

	Call ExtractDateFrom(frm1.txtFrYyyymm.Text,frm1.txtFrYyyymm.UserDefinedFormat,Parent.gComDateType,strFrYear,strFrMonth,strFrDay)
	Call ExtractDateFrom(frm1.txtToYyyymm.Text,frm1.txtToYyyymm.UserDefinedFormat,Parent.gComDateType,strToYear,strToMonth,strToDay)

	strTrFrDt		= strFrYear & strFrMonth & strFrDay
	strTrToDt		= strToYear & strToMonth & strToDay

	strUrl	= strUrl & "tr_fr_dt|"		& strTrFrDt
	strUrl	= strUrl & "|tr_to_dt|"		& strTrToDt

	SetPrintCond = True
End Function

Function FncBtnPrint() 

    Dim StrEbrFile, strUrl
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPrint(EBAction,ObjName,strUrl)	

End Function

Function FncBtnPreview() 
    
    Dim StrEbrFile, strUrl

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebc")
	call FncEBCPreview(ObjName , strUrl)
	
End Function

Function FncPrint() 
    Call parent.FncPrint()
End Function

Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLE,False)
End Function

Function FncExit()
    FncExit = True
End Function

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
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>데이타현황</font></td>
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>모듈</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ModuleOpt" CHECKED ID="ModuleOpt2" VALUE="N" tag="25"><LABEL FOR="ModuleOpt2">영업</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ModuleOpt" ID="ModuleOpt3" VALUE="N" tag="25"><LABEL FOR="ModuleOpt3">생산</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ModuleOpt" ID="ModuleOpt1" VALUE="Y" tag="25"><LABEL FOR="ModuleOpt1">구매</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ModuleOpt" ID="ModuleOpt4" VALUE="N" tag="25"><LABEL FOR="ModuleOpt4">품질</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ModuleOpt" ID="ModuleOpt5" VALUE="N" tag="25"><LABEL FOR="ModuleOpt5">회계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ModuleOpt" ID="ModuleOpt6" VALUE="N" tag="25"><LABEL FOR="ModuleOpt6">인사</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">발생</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">입력</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>기간</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/b7000oa1_fpDateTime1_txtFrYyyymm.js'></script>&nbsp;~&nbsp;
														<script language =javascript src='./js/b7000oa1_fpDateTime2_txtToYyyymm.js'></script></TD>
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
						<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
<!--						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>	-->
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="strUrl" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

