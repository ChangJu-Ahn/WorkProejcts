<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4512QA1
'*  4. Program Name         : 출하상세조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Date표준적용 
'*                            2002/12/16 Include 성능향상 강준구 
'*                            -2002/12/20 : Get방식 을 Post방식으로 변경 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="S4512qa1.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, Parent.gDateFormat)

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

</SCRIPT>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출하요청상세</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*">&nbsp;</td>
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
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6><INPUT NAME="txtconBp_cd" ALT="납품처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtDNType" SIZE=10 MAXLENGTH=3 TAG="11XXXU" ALT="출하형태" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSORef" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtDNTypeNm" SIZE=20 TAG="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>	
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="공장" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" Onclick="vbscript:OpenConSItemDC 4">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>창고</TD>
									<TD CLASS=TD6><INPUT NAME="txtStoRo_cd" ALT="창고" TYPE="Text" MAXLENGTH=7 SiZE=10 tag="11XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 5">&nbsp;<INPUT NAME="txtStoRo_Nm" TYPE="Text" SIZE=20 tag="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14" CLASS="protected" READONLY="true" TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>출하등록잔량</TD> 
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="A" CHECKED ID="rdoQueryFlg1"><LABEL FOR="rdoQueryFlg1">전체</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="Y" CHECKED ID="rdoQueryFlg2"><LABEL FOR="rdoQueryFlg2">없음</LABEL>&nbsp;&nbsp;&nbsp;
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryFlg" TAG="11X" VALUE="N" ID="rdoQueryFlg3"><LABEL FOR="rdoQueryFlg3">있음</LABEL>			
									</TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 id="lblTitle" NOWRAP>출고예정일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtPromiseFrDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="조회시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtPromiseToDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="조회종료일"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 6"></TD>	
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="OPMD_UMODE" tag="24">

<INPUT TYPE=HIDDEN NAME="HtxtconBp_cd" tag="24"> 
<INPUT TYPE=HIDDEN NAME="HtxtDNType" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtStoRo_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtPromiseFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HtxtPromiseToDt" tag="24">	    
<INPUT TYPE=HIDDEN NAME="HtxtRadio" tag="24">	    
<INPUT TYPE=HIDDEN NAME="HtxtTrackingNo" tag="24">   

<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14">	

<INPUT TYPE=HIDDEN NAME="txt_lgPageNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgStrPrevKey" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgMaxCount" tag="24" TABINDEX="-1">  
<INPUT TYPE=HIDDEN NAME="txt_lgSelectListDT" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgTailList" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txt_lgSelectList" tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
