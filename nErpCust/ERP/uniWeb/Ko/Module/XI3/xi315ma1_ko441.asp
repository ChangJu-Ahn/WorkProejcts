<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : MES
'*  2. Function Name        : 출하관리
'*  3. Program ID           : xi315MA1_KO441
'*  4. Program Name         : 제품재고현황
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="xi315MA1_KO441.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
Dim EndDate
Dim strYear, strMonth, strDay

	iDBSYSDate = "<%=GetSvrDate%>"											'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'	StartDate = UNIDateAdd("D",-1,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 처음 날짜 
	Call	ExtractDateFrom(iDBSYSDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)
	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UNIDateAdd("D", 0,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜 

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "*", "NOCOOKIE", "QA") %>	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>제품재고현황(S)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>MES송신기간</TD>
								    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtSendStartDt 	CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="11X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtSendEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="11X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>SEC 품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=text NAME=txtSecItemCd SIZE=15 MAXLENGTH=18 tag=11xxxU ALT="SEC품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME=btnItemCd ALIGN=top TYPE=button ONCLICK="vbScript:Call OpenSecItem">&nbsp;<INPUT TYPE=text NAME=txtSecItemNm SIZE=20 tag=14 TABINDEX=-1></TD>
									<TD CLASS=TD5 NOWRAP>생산일자</TD>
									<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtPlanStartDt	CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="시작일" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									&nbsp;~&nbsp;
									<SCRIPT LANGUAGE=JAVASCRIPT>
									ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> name=txtPlanEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="종료일" tag="12X1" id=fpDateTime1></OBJECT>');
									</SCRIPT>
									</TD>
								</TR>								
								<TR>									
									<TD CLASS=TD5 NOWRAP>구분</TD>
									<TD CLASS=TD6 colspan =3><INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunMU	TAG="12" CHECKED onclick=radio1_onchange()><LABEL FOR=rdoGubunMU>양산가용재고</LABEL>
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunSU	TAG="12" onclick=radio2_onchange()><LABEL FOR=rdoGubunSU>샘플가용재고</LABEL>
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunMH TAG="12" onclick=radio3_onchange()><LABEL FOR=rdoGubunMH>양산재고(H)</LABEL>
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunSH TAG="12" onclick=radio4_onchange()><LABEL FOR=rdoGubunSH>샘플재고(H)</LABEL>
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunMInv TAG="12" onclick=radio5_onchange()><LABEL FOR=rdoGubunMInv>양산입고</LABEL>									  
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunSInv TAG="12" onclick=radio6_onchange()><LABEL FOR=rdoGubunSInv>샘플입고</LABEL>
							<!--		</TD>												  
								</TR>
								<TR>									
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 colspan =3>--><INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunMOut TAG="12" onclick=radio7_onchange()><LABEL FOR=rdoGubunMOut>양산출고</LABEL> 
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunSOut TAG="12" onclick=radio8_onchange()><LABEL FOR=rdoGubunSOut>샘플출고</LABEL>	
												  <INPUT TYPE=radio CLASS=Radio NAME=rdoGubun ID=rdoGubunVOut TAG="12" onclick=radio9_onchange()><LABEL FOR=rdoGubunVOut>가상출고</LABEL>	
									</TD>												  
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD COLSPAN=4 CLASS=TD5 NOWRAP>
									<TABLE CELLSPACING=3 CELLPADDING=0 BORDER=0>
										<TR>
											<TD>양산입고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtMassSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>샘플입고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtSampleSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>입고계</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtInventorySumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>양산출고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtMOutSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>샘플출고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtSampleOutSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>가상출고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtVOutSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>출고계</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtOutSumQty.js"></SCRIPT>&nbsp;</TD>											
										</TR>
											<TD>양산재고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtMGoodsSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>샘플재고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtSampleGoodsSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>재고계</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtGoodsSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>양산재고(H)</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtMHoldSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>샘플재고(H)</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtSampleHoldSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>양산가용재고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtMUseSumQty.js"></SCRIPT>&nbsp;</TD>
											<TD>샘플가용재고</TD>
											<TD><SCRIPT LANGUAGE=javaScript SRC="./js/xi315ma1_KO441_txtSampleUseSumQty.js"></SCRIPT>&nbsp;</TD>																																								
										<TR>
										</TR>
									</TABLE>
								</TD>
							</TR>	
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
		            FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txthPlanStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txthPlanEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txthSendStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txthSendEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txthSecItemCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
