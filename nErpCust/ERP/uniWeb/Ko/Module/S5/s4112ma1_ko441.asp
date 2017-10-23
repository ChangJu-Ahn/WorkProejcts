<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4112MA1
'*  4. Program Name         : 출하내역등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41111MaintDnHdrSvr, S41121MaintDnDtlSvr, S41115PostGoodsIssueSvr
'*         S14113ChkDnCreditLimitSvr, S14114ChkGiCreditLimitSvr   
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/12/23 : include 성능향상만 반영 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/JpQuery.vbs">				</SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="S4112ma1_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                               

Dim iDBSYSDate
Dim EndDate, StartDate

Dim lblnWinEvent   '박정순 추가 
Dim interface_Account   '박정순 추가 


iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'=========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
<% Call LoadBNumericFormatA( "I", "*", "NOCOOKIE", "MA") %>
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출하내역등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenSODtlRef1">수주내역등록(상품&원재료)</A>&nbsp&nbsp&nbsp;<A href="vbscript:OpenSODtlRef">수주내역및거래처재고참조</A>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>출하번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConDnNo" ALT="출하번호" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDnDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnDtl()"></TD>
									<TD CLASS="TDT"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6"><INPUT NAME="txtPlantCd" ALT="공장" TYPE="Text" SiZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								<TD CLASS="TD5" NOWRAP>납품처</TD>
								<TD CLASS="TD6"><INPUT NAME="txtShipToParty" ALT="납품처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtShipToPartyNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>       
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>출하형태</TD>
								<TD CLASS="TD6"><INPUT NAME="txtDnType" ALT="출하형태" TYPE="Text" MAXLENGTH=3 SiZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtDnTypeNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>       
								<TD CLASS="TD5" NOWRAP>수주형태</TD>
								<TD CLASS="TD6"><INPUT NAME="txtSoType" ALT="수주형태" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtSoTypeNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>       
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>출고예정일</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlannedGIDt" ALT="납기일" TYPE="Text" style="text-align=center" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								<TD CLASS=TD5 NOWRAP>납기일</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDlvyDt" ALT="납기일" TYPE="Text" style="text-align=center" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
							</TR>
							<TR> 
								<TD CLASS="TD5" NOWRAP>수주번호</TD>
								<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								<TD CLASS="TD5" NOWRAP>합계(picking수량)</TD>        
								<TD CLASS="TD6"><INPUT NAME="txtSumPicking" ALT="picking수량" TYPE="Text" MAXLENGTH=18 SiZE=20 style="text-align=right" tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1"></TD> 
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100% CLASS=TD6 COLSPAN=4>
									<FIELDSET ID="filPost" CLASS="CLSFLD" TITLE="출고처리">
									<LEGEND ALIGN=LEFT>출고처리</LEGEND>
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP STYLE="PADDING-BOTTOM:5px">실제출고일</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtActualGIDt" CLASS=FPDTYYYYMMDD tag="24X1" ALT="실제출고일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
												</TD>
												<TD CLASS=TD5 NOWRAP>재고담당자</TD>
												<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvMgr" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="재고담당자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInvMgr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInvMgrPopUp">&nbsp;<INPUT NAME="txtInvMgrNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="재고담당자명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>후속작업여부</TD>
												<TD CLASS=TD6 NOWRAP>
													<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="24" READONLY="true" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">매출채권</LABEL>&nbsp;&nbsp;
													<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="24" READONLY="true" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">세금계산서</LABEL>
												</TD>
												<TD CLASS="TD5" NOWRAP>출고번호</TD>
												<TD CLASS="TD6"><INPUT NAME="txtGINo" ALT="출고번호" TYPE="Text" SiZE=20 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnPosting" CLASS="CLSMBTN">출고처리</BUTTON>&nbsp;
						<BUTTON NAME="btnPostCancel" CLASS="CLSMBTN">출고처리취소</BUTTON>&nbsp;
			         		<BUTTON NAME="btnGlSel" CLASS="CLSSBTN"  ONCLICK="OpenGlRef()">전표조회</BUTTON>
					</TD>
					<TD WIDTH=* Align=Right><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">출하등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadIns tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadUpd tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadDel tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtHDnNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHRetFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtArFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtVatFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtRetBillFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtExportFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHRefRoot" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
   <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
  </DIV>
</BODY>
</HTML>
