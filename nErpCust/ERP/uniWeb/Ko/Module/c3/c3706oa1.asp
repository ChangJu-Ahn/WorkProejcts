
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3706oa1
'*  4. Program Name         : 실제원가 출력 
'*  5. Program Desc         : 실제원가 출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/01/15
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Hyo Seok, Seo
'* 10. Modifier (Last)      : Cho Ig sung
'* 11. Comment              :
'=======================================================================================================  -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit


Dim lgBlnFlgChgValue
Dim lgIntFlgMode
Dim lgIntGrpCount

Dim IsOpenPop          

Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
         
End Sub

Sub SetDefaultVal()
	Dim StartDate

	StartDate	= "<%=GetSvrDate%>"
'	EndDate		= UNIDateAdd("m", -1, StartDate,Parent.gServerDateFormat)

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "PA") %>
End Sub

Function OpenPopup(Byval param, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
		Case 1
			arrParam(0) = "품목계정 팝업"
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"
			arrParam(2) = Trim(frm1.txtItmeAcct.value)
			arrParam(3) = ""
			arrParam(4) = "a.MAJOR_CD = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group <> '6MRO' "
			arrParam(5) = "품목계정"
	
			arrField(0) = "minor_CD"
			arrField(1) = "minor_NM"
    
			arrHeader(0) = "품목계정코드"
			arrHeader(1) = "품목계정명"
		Case 2
			arrParam(0) = "작업단계팝업"
			arrParam(1) = "B_MINOR M, B_CONFIGURATION C"
			arrParam(2) = Trim(frm1.txtWorkStepCd.Value)
			arrParam(3) = ""
			arrParam(4) = "M.MINOR_CD = C.MINOR_CD and M.MAJOR_CD = C.MAJOR_CD and C.SEQ_NO = 4 and C.REFERENCE = " & FilterVar("Y", "''", "S") & "  and M.MAJOR_CD = " & FilterVar("C2000", "''", "S") & " "
			arrParam(5) = "작업단계"
	
			arrField(0) = "M.minor_CD"
			arrField(1) = "M.minor_NM"
    
			arrHeader(0) = "작업단계코드"
			arrHeader(1) = "작업단계명"
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	  If iWhere = 1 Then
		frm1.txtItmeAcct.focus
	  ElseIf iWhere = 2 Then
	    frm1.txtWorkStepCd.focus
      End If	    	
		Exit Function
	Else
		Call SetReturnPopUp(arrRet, iWhere)
	End If
		
End Function

Function SetReturnPopUp(ByVal arrRet, ByVal iWhere)
	
	Select Case iWhere
		Case 1
			frm1.txtItmeAcct.focus
			frm1.txtItmeAcct.value		= arrRet(0)		
			frm1.txtItemAcctNm.value	= arrRet(1)		
		Case 2
			frm1.txtWorkStepCd.focus
			frm1.txtWorkStepCd.Value	= arrRet(0)		
			frm1.txtWorkStepNm.Value	= arrRet(1)		
		Case Else
			Exit Function
	End select	

End Function

Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitVariables                                        
    Call SetDefaultVal
    Call SetToolbar("10000000000011")
    
    frm1.txtYyyymm.focus 
    Set gActiveElement = document.activeElement	
    
End Sub

Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

Sub txtYyyymm_DblClick(Button)
	If Button = 1 Then
		frm1.txtYyyymm.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtYyyymm.focus
	End If
End Sub


Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

Function FncQuery() 
    FncBtnPreview()
End Function

Sub SetPrintCond(StrEbrFile, stryyyymm, strworkstep, stritemacct)
Dim	strYear, strMonth, strDay

	if frm1.PrintOpt1.checked = True then
		StrEbrFile = "c3706oa1"
	elseif frm1.PrintOpt2.checked = True then
		StrEbrFile = "c3706oa2"
	elseif frm1.PrintOpt3.checked = True then
		StrEbrFile = "c3706oa3"
	elseif frm1.PrintOpt4.checked = True then
		StrEbrFile = "c3706oa4"
	elseif frm1.PrintOpt5.checked = True then
		StrEbrFile = "c3706oa5"
	else
		StrEbrFile = "c3706oa6"
	end if

	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	stryyyymm	= strYear & strMonth
	stritemacct	= Trim(UCase(frm1.txtItmeAcct.value))
	strworkstep	= Trim(UCase(frm1.txtWorkStepCd.value))
	
	if stritemacct = "" then
		stritemacct = "%"
		frm1.txtItemAcctNm.value	= ""
	End if	

	if strworkstep = "" then
		strworkstep = "%"
		frm1.txtWorkStepNm.value = ""
	End if	

End Sub

Function FncBtnPrint() 
 
    Dim StrEbrFile
	Dim condvar
	dim stryyyymm, strworkstep, stritemacct

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, stryyyymm, strworkstep, stritemacct)

	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & stryyyymm
	condvar	= condvar & "|itemacct|" & stritemacct
	if frm1.PrintOpt1.checked = True or frm1.PrintOpt2.checked = True or frm1.PrintOpt3.checked = True then
		condvar	= condvar & "|workstep|" & strworkstep
	end if
	
	call FncEBRPrint(EBAction,ObjName,condvar)	    
    
End Function

Function FncBtnPreview() 
    
    Dim StrEbrFile
	Dim condvar
	dim stryyyymm, strworkstep, stritemacct

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, stryyyymm, strworkstep, stritemacct)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & stryyyymm
	condvar	= condvar & "|itemacct|" & stritemacct
	if frm1.PrintOpt1.checked = True or frm1.PrintOpt2.checked = True or frm1.PrintOpt3.checked = True then
		condvar	= condvar & "|workstep|" & strworkstep
	end if
					
	call FncEBRPreview(ObjName,condvar)
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실제원가출력</font></td>
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
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">상세</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">집계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt3" VALUE="N" tag="25"><LABEL FOR="PrintOpt3">작업단계</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt4" VALUE="N" tag="25"><LABEL FOR="PrintOpt4">구성비율</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt5" VALUE="Y" tag="25"><LABEL FOR="PrintOpt5">입고차이</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt6" VALUE="N" tag="25"><LABEL FOR="PrintOpt6">재고차이</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/c3706oa1_fpDateTime1_txtYyyymm.js'></script>
								</TD>								
							</TR>
							<TR>	
								<TD CLASS="TD5">품목계정</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtItmeAcct" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtItmeAcct.value,1)">
									 <INPUT TYPE=TEXT ID="txtItemAcctNm" NAME="txtItemAcctNm" SIZE=30 tag="14X">
								</TD>
							</TR>
							<TR>	
								<TD CLASS="TD5">작업단계</TD>
								<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtWorkStepCd" SIZE=10 MAXLENGTH=2 tag="11XXXU" ALT="작업단계"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWorkStepCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtWorkStepCd.value,2)">
									 <INPUT TYPE=TEXT ID="txtWorkStepNm" NAME="txtWorkStepNm" SIZE=30 tag="14X">
								</TD>
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
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>
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
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >	
</FORM>
</BODY>
</HTML>

