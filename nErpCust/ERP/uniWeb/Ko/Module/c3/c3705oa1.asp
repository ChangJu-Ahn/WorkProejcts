
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3705oa1
'*  4. Program Name         : 재공배부내역출력 
'*  5. Program Desc         : 재공배부내역출력 
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

Function OpenPopup(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	

	Dim strYear,strMonth,strDay,strYYYYMM


	stryyyymm = strYear & strMonth


	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	if iWhere = 0 then

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

	elseif iWhere = 1 then
		Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

		stryyyymm = strYear & strMonth
	
		arrParam(0) = "코스트센타팝업"
		arrParam(1) = "B_COST_CENTER a, C_WORKIN_MFC_ALLC b"
		arrParam(2) = strCode
		arrParam(3) = ""
		arrParam(4) = "a.COST_CD = b.COST_CD and b.prnt_bas_amt <> 0 and b.YYYYMM =  " & FilterVar(stryyyymm , "''", "S") & ""
		arrParam(5) = "코스트센타"			
	
		arrField(0) = "a.COST_CD"
		arrField(1) = "a.COST_NM"
    
		arrHeader(0) = "코스트센타코드"
		arrHeader(1) = "코스트센타명"
	else
		Exit Function
	end if
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	 If iWhere = 0 Then
		frm1.txtPlantCd.focus
	 ElseIf iWhere = 1 Then
		frm1.txtCostCd.focus
	 End If
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If	

End Function

Function SetPopup(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
			.txtPlantCd.focus
    		.txtPlantCd.value = arrRet(0)
    		.txtPlantNm.value = arrRet(1)
		ElseIf iWhere = 1 Then
			.txtCostCd.focus
    		.txtCostCd.value = arrRet(0)
    		.txtCostNm.value = arrRet(1)
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

Sub SetPrintCond(StrEbrFile, stryyyymm, strPlantCd, strCostCd)
Dim	strYear, strMonth, strDay

	StrEbrFile = "c3705oa1"

	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyyMm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	stryyyymm = strYear & strMonth

	strPlantCd = Trim(UCase(frm1.txtPlantCd.value))
	
	if strPlantCd = "" then
		strPlantCd = "%"
		frm1.txtPlantNm.value = ""
	End if	
	strCostCd = Trim(UCase(frm1.txtCostCd.value))
	
	if strCostCd = "" then
		strCostCd = "%"
		frm1.txtCostNm.value = ""
	End if	
	
End Sub

Function FncBtnPrint() 
'	Dim strUrl
'	Dim lngPos
'	Dim intCnt
    Dim StrEbrFile
	dim stryyyymm,strPlantCd,strCostCd
	Dim condvar
	
    If Not chkField(Document, "1") Then
       Exit Function
    End If

 
	Call SetPrintCond(StrEbrFile, stryyyymm, strPlantCd,strCostCd)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & stryyyymm
	condvar	= condvar & "|plantcd|" & strPlantCd
	condvar	= condvar & "|costcd|" & strCostCd

	call FncEBRPrint(EBAction,ObjName,condvar)	
    
End Function

Function FncBtnPreview() 

    Dim StrEbrFile
	dim stryyyymm,strCostCd,strPlantCd
	Dim condvar
	
    If Not chkField(Document, "1") Then
       Exit Function
    End If

	Call SetPrintCond(StrEbrFile, stryyyymm, strPlantCd,strCostCd)
	
	ObjName = AskEBDocumentName(StrEBrFile, "ebr")

'   표준변경 적용(2002.1.14)	
	condvar = condvar & "yyyymm|" & stryyyymm
	condvar	= condvar & "|plantcd|" & strPlantCd
	condvar	= condvar & "|costcd|" & strCostCd

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
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재공배부내역출력</font></td>
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
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업년월</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/c3705oa1_fpDateTime1_txtYyyymm.js'></script>
									</TD>								
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtPlantCd.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
								</TR>								
								<TR>	
									<TD CLASS="TD5" NOWRAP>코스트센타</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="코스트센타"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup frm1.txtCostCd.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtCostNm" SIZE=30 tag="14"></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
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

