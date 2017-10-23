<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 매출채권현황출력 
'*  3. Program ID           : s5111oa2
'*  4. Program Name         : 매출채권현황출력 
'*  5. Program Desc         : 매출채권현황출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/18
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Cho Sung Hyun
'* 10. Modifier (Last)      : Hwang Seongbae
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

Option Explicit

Const C_PopSoldToParty	= 1
Const C_PopSalesGrp		= 2

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size

Dim IsOpenPop          

'=========================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'=========================================
Sub SetDefaultVal()

	frm1.txtBillFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtBillToDt.Text = EndDate
	frm1.txtBillFromDt.focus 
	
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "OA") %>	
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	OpenConPopup = False
	IsOpenPop = True

	With frm1
		Select Case pvIntWhere
			' 주문처 
			Case C_PopSoldToParty												
				iArrParam(1) = "B_BIZ_PARTNER"
				iArrParam(2) = Trim(.txtSoldToParty.value)
				iArrParam(3) = ""
'				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"
				iArrParam(4) = "BP_TYPE IN (" & FilterVar("CS", "''", "S") & ", " & FilterVar("C", "''", "S") & " )"
				iArrParam(5) = .txtSoldToParty.Alt
					
				iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"
				    
				iArrHeader(0) = .txtSoldToParty.Alt
				iArrHeader(1) = .txtSoldToPartyNm.Alt

				.txtSoldToParty.focus

			' 영업그룹 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(.txtSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				iArrParam(5) = .txtSalesGrp.Alt
				
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtSalesGrp.Alt
			    iArrHeader(1) = .txtSalesGrpNm.Alt
			    
			    .txtSalesGrp.focus

		End Select
	End With

	iArrParam(0) = iArrParam(5)							' 팝업 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) <> "" Then OpenConPopup = SetConPopup(iArrRet, pvIntWhere)

End Function

'=========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1	
		Select Case pvIntWhere
			Case C_PopSoldToParty
				.txtSoldToParty.Value	= pvArrRet(0)
				.txtSoldToPartyNm.Value	= pvArrRet(1)

			' 영업그룹 
			Case C_PopSalesGrp
				.txtSalesGrp.value = pvArrRet(0) 
				.txtSalesGrpNm.value = pvArrRet(1)
		End Select
	End With
	
	SetConPopup = True
End Function

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
Sub txtBillFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBillFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillFromDt.Focus
    End If
End Sub

'==========================================
Sub txtBillToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBillToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillToDt.Focus
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

	If ValidDateCheck(frm1.txtBillFromDt, frm1.txtBillToDt) = False Then Exit Function
    
	If UCase(frm1.txtSoldToParty.value) = "" Then
		iStrUrl = "BillToParty|%"
	Else
		iStrUrl = "BillToParty|" & Replace(Trim(frm1.txtSoldToParty.value), "'" ,  "''")
	End If

	iStrUrl = iStrUrl & "|BillFromDt|" & UniConvDateToYYYYMMDD(frm1.txtBillFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	iStrUrl = iStrUrl & "|BillToDt|" & UniConvDateToYYYYMMDD(frm1.txtBillToDt.Text,Parent.gDateFormat,Parent.gServerDateType)

	If UCase(frm1.txtSalesGrp.value) = "" Then
		iStrUrl = iStrUrl & "|SalesGrp|%"
	Else
		iStrUrl = iStrUrl & "|SalesGrp|" & Replace(Trim(frm1.txtSalesGrp.value), "'" ,  "''")
	End If

	ObjName = AskEBDocumentName("s5111oa2","ebr")
	
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


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권현황</font></td>
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
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtBillFromDt" CLASS="FPDTYYYYMMDD" tag="12X1" ALT="매출채권시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
											<TD>
												&nbsp;~&nbsp;
											</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtBillToDt" CLASS="FPDTYYYYMMDD" tag="12X1" ALT="매출채권종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>주문처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldToParty" ALT="주문처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoldToParty" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPopup C_PopSoldToParty">&nbsp;<INPUT NAME="txtSoldToPartyNm" ALT="주문처명" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK ="OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=25 TAG="14" ALT="영업그룹명"></TD>
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
