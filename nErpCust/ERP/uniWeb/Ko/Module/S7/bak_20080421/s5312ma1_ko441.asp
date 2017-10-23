<%@ LANGUAGE="VBSCRIPT" %>

<%'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5312MA1
'*  4. Program Name         : 세금계산서내역등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : S53119LookupTaxBillHdrSvr, S53128ListTaxBillDtlSvr, S53121MaintTaxBillDtlSvr, S53115PostTaxBillSvr
'*  7. Modified date(First) : 2001/06/26
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2001/06/26 : 6차 화면 layout & ASP Coding
'*                            -2001/11/09 : 부가세별로 계산하는 로직 추가 
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="S5312ma1_ko441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgvariables.inc" -->	
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서내역등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenBillDtlRef">매출채권내역참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>세금계산서관리번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTaxBillNo" ALT="세금계산서관리번호" TYPE="Text" MAXLENGTH=18 SiZE=30 tag="12XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBillNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTaxBillNo()"></TD>
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
								<TD CLASS=TD5 NOWRAP>발행처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBilltoParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" ALT="발행처" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtBilltoPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								<TD CLASS=TD5 NOWRAP>매출채권번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillNo" ALT="매출채권번호" TYPE="Text" MAXLENGTH="18" SIZE=30 tag="24XXXU" class = protected readonly = true TABINDEX="-1"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT적용기준</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoVatCalcType" id="rdoVatCalcType1" value="1" tag = "24">
										<label ID="lblVatCalcType1" for="rdoVatCalcType1">개별</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoVatCalcType" id="rdoVatCalcType2" value="2" tag = "24" checked>
										<label ID="lblVatCalcType2" for="rdoVatCalcType2">통합</label>
								</TD>
								<TD CLASS=TD5 NOWRAP>VAT포함구분</TD>
								<TD CLASS=TD6 NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoVatIncFlag" id="rdoVatIncFlag1" value="1" tag = "24">
										<label ID="lblVatIncFlag1" for="rdoVatIncFlag1">별도</label>&nbsp;&nbsp;&nbsp;&nbsp;
									<input type=radio CLASS = "RADIO" name="rdoVatIncFlag" id="rdoVatIncFlag2" value="2" tag = "24" checked>
										<label ID="lblVatIncFlag2" for="rdoVatIncFlag2">포함</label>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT유형</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 ALT="VAT유형" tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="25" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								<TD CLASS=TD5 NOWRAP>VAT율</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5312ma1_fpDoubleSingle1_txtVatRate.js'></script>&nbsp;<LABEL><b>%</b></LABEL>
											</TD>
											
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>공급가액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5312ma1_fpDoubleSingle2_txtSupplyAmt.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" class = protected readonly = true TABINDEX="-1">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS="TD5" NOWRAP>공급가자국액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/s5312ma1_fpDoubleSingle3_txtSupplyLocAmt.js'></script>
											</TD>
											<TD>
												&nbsp;<INPUT NAME="txtLocCur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" class = protected readonly = true TABINDEX="-1">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5312ma1_fpDoubleSingle4_txtVatAmt.js'></script>							
								</TD>
								<TD CLASS=TD5 NOWRAP>VAT자국금액</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5312ma1_fpDoubleSingle5_txtLocVatAmt.js'></script>							
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<script language =javascript src='./js/s5312ma1_OBJECT1_vspdData.js'></script>
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
			<TABLE <%=LR_SPACE_TYPE_30%>border =1>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPostFlag" CLASS="CLSMBTN">발행</BUTTON></TD>
					<TD WIDTH=* Align=Right><a href = "vbscript:JumpChgCheck(BIZ_BillTax_JUMP_ID)">세금계산서등록</a></TD>
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="HTaxBillNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HXchRate" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HXchRateOp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HPostFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HSalesGrpCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="HSalesGrpNm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtIssueDt" tag="24" TABINDEX="-1">

</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
