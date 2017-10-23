
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3703oa1
'*  4. Program Name         : 실제 BOM 출력 
'*  5. Program Desc         : 실제 BOM 출력 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2001/01/15
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Hyo Seok, Seo
'* 10. Modifier (Last)      : Cho Ig sung
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

Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Dim strYear,strMonth,strDay,strYYYYMM
	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	stryyyymm = strYear & strMonth


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	if iWhere = 0 then
		arrParam(0) = "품목팝업"
		arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "a.ITEM_CD = b.item_cd and (exists (select top 1 * from c_bom_rcpt where prnt_plant_cd = b.plant_cd and prnt_item_cd = b.item_cd and rcpt_amt <> 0 and yyyymm= " & FilterVar(stryyyymm, "''", "S") & ")"
		arrParam(4) = arrParam(4) & " or exists (select top 1 * from c_bom_issue where child_plant_cd = b.plant_cd and child_item_cd = b.item_cd and (bas_amt+issue_amt) <> 0 and yyyymm= " & FilterVar(stryyyymm, "''", "S") & "))"
		IF frm1.txtPlantCd.value <> "" Then 
			arrParam(4) = arrParam(4) & " and b.plant_cd = " & FilterVar(frm1.txtPlantCd.value, "''", "S")
		END IF


		IF frm1.txtItemAcct.value  <> "" Then 
			arrParam(4) = arrParam(4) & " and b.item_acct = " & FilterVar(frm1.txtItemAcct.value, "''", "S")
		END IF

		arrParam(5) = "품목"			
	
		arrField(0) = "a.ITEM_CD"
		arrField(1) = "a.ITEM_NM"
		 
		arrHeader(0) = "품목코드"
		arrHeader(1) = "품목명"
	elseif iWhere = 1 then
		arrParam(0) = "공장팝업"
		arrParam(1) = "B_PLANT"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = ""
		arrParam(5) = "공장"			
	
		arrField(0) = "PLANT_CD"
		arrField(1) = "PLANT_NM"
		 
		arrHeader(0) = "공장코드"
		arrHeader(1) = "공장명"

	elseif iWhere = 2 then
		arrParam(0) = "품목계정팝업"
		arrParam(1) = "B_MINOR a,b_item_acct_inf b"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  and a.minor_cd = b.item_acct and b.item_acct_group <> " & FilterVar("6MRO","''","S")
		arrParam(5) = "품목계정"			
	
		arrField(0) = "MINOR_CD"
		arrField(1) = "MINOR_NM"
		 
		arrHeader(0) = "품목계정"
		arrHeader(1) = "품목계정명"
	else
		Exit Function
	end if
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	 If iWhere = 0 Then
		frm1.txtItemCd.focus
	 ElseIf iWhere = 1 Then
		frm1.txtPlantCd.focus
	 ElseIf iWhere = 2 Then
		frm1.txtItemAcct.focus
	 End If
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

Function SetPopup(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		.txtItemCd.focus
    		.txtItemCd.value = arrRet(0)
    		.txtItemNm.value = arrRet(1)
		ElseIf iWhere = 1 Then
			.txtPlantCd.focus
    		.txtPlantCd.value = arrRet(0)
    		.txtPlantNm.value = arrRet(1)
		ElseIf iWhere = 2 Then
			.txtItemAcct.focus
    		.txtItemAcct.value = arrRet(0)
    		.txtItemAcctNm.value = arrRet(1)
    	Else
			Exit Function
    	End If
	End With
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

Sub SetPrintCond(StrEbrFile, stryyyymm, strPlantCd, strItemCd, strItemAcct)
Dim	strYear, strMonth, strDay

	if frm1.PrintOpt1.checked = True then
		StrEbrFile = "c3703oa1"
	elseif frm1.PrintOpt2.checked = True then
		StrEbrFile = "c3703oa2"
	elseif frm1.PrintOpt3.checked = True then
		StrEbrFile = "c3703oa3"
	elseif frm1.PrintOpt4.checked = True then
		StrEbrFile = "c3703oa4"
	elseif frm1.PrintOpt5.checked = True then
		StrEbrFile = "c3703oa5"
	elseif frm1.PrintOpt6.checked = True then
		StrEbrFile = "c3703oa6"
	else 
		Exit Sub
	end if

	Call ExtractDateFrom(frm1.txtYyyyMm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	stryyyymm = strYear & strMonth
	strPlantCd = Trim(UCase(frm1.txtPlantCd.value))
	strItemCd = Trim(UCase(frm1.txtItemCd.value))
	strItemAcct = Trim(UCase(frm1.txtItemAcct.value))
	
	if strPlantCd = "" then
		strPlantCd = "%"
		frm1.txtPlantNm.value = ""
	End if	

	if strItemCd = "" then
		strItemCd = "%"
		frm1.txtItemNm.value = ""
	End if	

	if strItemAcct = "" then
		strItemAcct = "%"
		frm1.txtItemAcctNm.value = ""
	End if	
End Sub

Function FncBtnPrint() 

    Dim StrEbrFile
    Dim condvar
	dim stryyyymm,strPlantCd,strItemCd,strItemAcct
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
	
	Call SetPrintCond(StrEbrFile, stryyyymm, strPlantCd, strItemCd, strItemAcct)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
	
	condvar = condvar & "yyyymm|" & stryyyymm
	condvar	= condvar & "|plantcd|" & strPlantCd
	condvar	= condvar & "|itemcd|" & strItemCd
	condvar	= condvar & "|ItemAcct|" & strItemAcct

	call FncEBRPrint(EBAction,ObjName,condvar)	
	 

End Function

Function FncBtnPreview() 
    
    Dim StrEbrFile
	Dim condvar
	dim stryyyymm,strPlantCd,strItemCd,strItemAcct

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, stryyyymm, strPlantCd, strItemCd, strItemAcct)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")
			
'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & stryyyymm
	condvar	= condvar & "|plantcd|" & strPlantCd
	condvar	= condvar & "|itemcd|" & strItemCd
	condvar	= condvar & "|ItemAcct|" & strItemAcct

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실제BOM출력</font></td>
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
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">입고집계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt5" VALUE="N" tag="25"><LABEL FOR="PrintOpt5">출고집계</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">상세</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt3" VALUE="N" tag="25"><LABEL FOR="PrintOpt3">입고</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt4" VALUE="N" tag="25"><LABEL FOR="PrintOpt4">출고</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt6" VALUE="N" tag="25"><LABEL FOR="PrintOpt6">요약</LABEL></SPAN>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/c3703oa1_fpDateTime1_txtYyyymm.js'></script></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtPlantCd.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>품목계정</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemAcct" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcct" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtItemAcct.value, 2">&nbsp;<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 tag="14"></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtItemCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
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

