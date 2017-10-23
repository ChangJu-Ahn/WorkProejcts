<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111MA8
'*  4. Program Name         : 출하현황조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41118ListDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ryu KYUNG RAE(1)
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO 변환 
'*                            -2002/12/16 : Include 성능향상 강준구 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="S4111ma8.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

</Script>

<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출하현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>
					<A href="vbscript:OpenSORef">출하내역현황</A>
					</TD>
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
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDn_Type" ALT="출하형태" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="11XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtDn_TypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_no" ALT="수주번호" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="11XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo()"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_to_party" ALT="납품처" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 2">&nbsp;<INPUT NAME="txtShip_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_meth" ALT="운송방법" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRequried 3">&nbsp;<INPUT NAME="txtTrans_meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU" STYLE="text-transform:uppercase" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSalesGrp()">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS="TD5" NOWRAP>출고예정일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s4111ma8_fpDateTime1_txtReqGiDtFrom.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s4111ma8_fpDateTime2_txtReqGiDtTo.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP>출고여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPostGiFlag" id="rdoPostGiFlagAll" value=" " tag = "11" checked>
											<label for="rdoPostGiFlagAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoPostGiFlag" id="rdoPostGiFlagYes" value="Y" tag = "11">
											<label for="rdoPostGiFlagYes">출고</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostGiFlag" id="rdoPostGiFlagNo" value="N" tag = "11">
											<label for="rdoPostGiFlagNo">미출고</label>
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/s4111ma8_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD WIDTH="*" Align=Right><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">출하등록</a></TD>
				<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		            FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPostGiFlag" tag="14">

<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHDn_Type" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSo_no" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHShip_to_party" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHReqGiDtFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHReqGiDtTo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHTrans_meth" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPostGiFlag" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
