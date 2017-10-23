<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111MA1
'*  4. Program Name         : 출하등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41111MaintDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2003/08/22
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
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
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="S4111ma1_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'=====================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=22>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="changeTabs(TAB1)">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출하등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="changeTabs(TAB2)">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납품 및 운송정보</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenSORef">수주참조</A>&nbsp;<A href="vbscript:OpenDNReqRef">출하요청참조</A></TD>
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
									<TD CLASS="TD6"><INPUT NAME="txtConDnNo" ALT="출하번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="12XXXU" CLASS="required" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConDnNo()"></TD>
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
					<TD WIDTH=100% VALIGN=TOP>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>출하번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtDnNo" ALT="출하번호" TYPE="Text" MAXLENGTH=18 SiZE=20 tag="25XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlantCd" ALT="공장" TYPE="Text" SIZE=10 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="공장명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출고예정일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtPlanned_gi_dt" CLASS=FPDTYYYYMMDD tag="22X1" ALT="출고예정일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMovType" ALT="출하형태" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtMovTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="22XXXU" CLASS="required" STYLE="text-transform:uppercase" ALT="납품처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShipToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup C_PopShiptoParty">&nbsp;<INPUT NAME="txtShip_to_partyNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="납품처명"></TD>
									<TD CLASS=TD5 NOWRAP>수주형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_Type" ALT="수주형태" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtSo_TypeNm" ALT="수주형태명" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>재고담당자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvMgr" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="재고담당자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInvMgr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup C_PopInvMgr">&nbsp;<INPUT NAME="txtInvMgrNm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="재고담당자명"></TD>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<INPUT NAME="txtSales_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_meth" ALT="운송방법" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransMeth" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPopup C_PopTransMeth">&nbsp;<INPUT NAME="txtTrans_meth_nm" TYPE="Text" SIZE=25 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1" ALT="운송방법명"></TD>
									<TD CLASS=TD5 NOWRAP>실제출고일</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtActGi_dt" ALT="실제출고일" TYPE="Text" style="text-align=center" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_no" ALT="수주번호" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkSoNo" tag="25" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkSoNo"> 수주번호지정</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>출고번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtGoods_mv_no" ALT="출고번호" TYPE="Text" MAXLENGTH="18" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5 NOWRAP>실제납품일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtArriv_dt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="실제납품일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>납기일</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDlvy_dt" ALT="납기일" TYPE="Text" style="text-align=center" SIZE=20 tag="24" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품시간</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtArriv_Tm" TYPE="Text" ALT="납품시간" MAXLENGTH="10" SIZE=40 tag="21"></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>							
								<TR>	
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="120" SIZE=91 ALT="비고" tag="21"></TD>
								</TR>
	                            <% Call SubFillRemBodyTD5656(7) %>
							</TABLE>
						</DIV>

						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처상세정보번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSTP_Inf_No" ALT="납품처상세정보번호" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<BUTTON NAME = "btnShipToPlceRef" CLASS="CLSMBTN">납품처상세정보참조</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>우편번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIP_cd" TYPE="Text" ALT="우편번호" MAXLENGTH="12" SIZE=20 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
									<TD CLASS=TD5 NOWRAP>인수자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiver" TYPE="Text" ALT="인수자명" MAXLENGTH="50" SIZE=35 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품주소</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1_Dlv" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2_Dlv" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR3_Dlv" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품장소</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtShip_to_place"  ALT="납품장소" TYPE="Text" MAXLENGTH="30" SIZE=91 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전화번호1</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No1" TYPE="Text" ALT="전화번호1" MAXLENGTH="20" SIZE=35 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
									<TD CLASS=TD5 NOWRAP>전화번호2</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTel_No2" TYPE="Text" ALT="전화번호2" MAXLENGTH="20" SIZE=35 tag="21XXXU" STYLE="text-transform:uppercase"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운송정보번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrnsp_Inf_No" ALT="운송정보번호" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="24XXXU" STYLE="text-transform:uppercase" CLASS="protected" READONLY="true" TABINDEX="-1">&nbsp;<BUTTON NAME = "btnTrnsMethRef" CLASS="CLSMBTN">운송정보참조</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>							
								<TR>
									<TD CLASS=TD5>운송회사</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtTransCo" SIZE=20 MAXLENGTH=50 TAG="21XXXX" ALT="운송회사"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransCo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTransCo()"></TD>
									<TD CLASS=TD5 NOWRAP>인계자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSender" TYPE="Text" ALT="인계자명" MAXLENGTH="50" SIZE=35 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5>차량번호</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 TAG="21XXXX" ALT="차량번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVehicleNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenVehicleNo()"></TD>
									<TD CLASS=TD5 NOWRAP>운전자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDriver" TYPE="Text" ALT="운전자명" MAXLENGTH="50" SIZE=35 tag="21"></TD>
								</TR>
	                            <% Call SubFillRemBodyTD5656(6) %>
							</TABLE>
						</DIV>
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
					<TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck()">출하내역등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtChkSoNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtTempSoNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHCntryCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtlgBlnChgValue1" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtlgBlnChgValue2" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHRefRoot" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
