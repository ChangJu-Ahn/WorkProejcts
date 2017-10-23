<%@ LANGUAGE="VBSCRIPT" %>
<!--
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 발주참조 Popup
'*  3. Program ID           : M3111RA3
'*  4. Program Name         : P/O Reference ASP
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/04/17
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/04 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/17 : ADO변환 
'**************************************************************************************
%>
-->
<HTML>
<HEAD>
<TITLE>입고참조</TITLE>
<!--
'########################################################################################################
'#						1. 선 언 부																		#
'########################################################################################################

'********************************************  1.1 Inc 선언  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncServer.asp" -->

<%
'============================================  1.1.1 Style Sheet  =======================================
%>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<%
'============================================  1.1.2 공통 Include  ======================================
%>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/eventpopup.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/AdoQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					<% '☜: indicates that All variables must be declared in advance %>
	

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
			

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>수혜자</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="수혜자"><IMG SRC="../../image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 0" >&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5>결제방법</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="결제방법"><IMG SRC="../../image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5>입고유형</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtMvmtType" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="입고유형"><IMG SRC="../../image/btnPopup.gif" NAME="btnMvmtType" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtMvmtTypeNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5>입고일</TD>
						<TD CLASS=TD6 NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m4444ra3_fpDateTime1_txtMVFrDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m4444ra3_fpDateTime2_txtMVToDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>	
						<TD CLASS=TD5 NOWRAP>구매그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="구매그룹"><IMG SRC="../../image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 3">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="24"></TD>
						<TD CLASS=TD5>발주형태</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPOType" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="발주형태"><IMG SRC="../../image/btnPopup.gif" NAME="btnPOType" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 4">&nbsp;<INPUT TYPE=TEXT NAME="txtPOTypeNm" SIZE=20 TAG="14"></TD>
					</TR>	
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m4444ra3_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=bizsize%>><IFRAME NAME="MyBizASP" SRC="m3111rb3.asp" WIDTH=100% HEIGHT=<%=bizsize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHBeneficiary" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHMvmtType" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHMVFrDt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHMVToDt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHPurGrp" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHPOType" TAG="14">
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</Form>
</BODY>
</HTML>