<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5311MA1
'*  4. Program Name         : 세금계산서등록 
'*  5. Program Desc         : 
'*  6. Comproxy List        : S53119LookupTaxBillHdrSvr, S53111MaintTaxBillHdrSvr
'*							  S53115PostTaxBillSvr, S51119LookupBillHdrSvr
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/03/27
'*                            2001/12/19	Date표준적용 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="s5311ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit
	
Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

<!-- #Include file="../../inc/lgvariables.inc" -->	

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>세금계산서</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>


					<TD WIDTH=* align=right><A href="vbscript:OpenBillRef">매출채권참조</A></TD>
	

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
									<TD CLASS=TD5 NOWRAP>세금계산서관리번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTaxbillNo" SIZE="20" MAXLENGTH="18" TAG="12XXXU" class=required STYLE="text-transform:uppercase" ALT="세금계산서관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxbillNo" ALIGN=top TYPE="BUTTON" OnClick="VBScript:OpenTaxbillNoPop()"></TD>
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>세금계산서관리번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxbillNo1" TYPE=TEXT SIZE="20" MAXLENGTH="18"  TAG="25XXXU" STYLE="text-transform:uppercase" ALT="세금계산서관리번호"></TD>
								<TD CLASS=TD5 NOWRAP>매출채권번호</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=TEXT NAME="txtBillNo" SIZE=20  MAXLENGTH=18 TAG="24XXXU" class = protected readonly = true TABINDEX="-1" ALT="매출채권번호">&nbsp;&nbsp;&nbsp;
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="25X" VALUE="Y" NAME="chkBillNoFlg" ID="chkBillNoFlg">
									<LABEL FOR="chkBillNoFlg">매출채권번호지정</LABEL>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>세금계산서번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxbillDocNo" TYPE=TEXT SIZE="30" MAXLENGTH="30"  TAG="21XXXU" class = protected readonly = true TABINDEX="-1" ALT="세금계산서번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBillDocNo" ALIGN=top TYPE="BUTTON" OnClick="VBScript:OpenTaxNo()" ></TD>
								<TD CLASS=TD5 NOWRAP>발행처</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBillToParty" SIZE="10" MAXLENGTH="10" TAG="24XXXU" class = protected readonly = true TABINDEX="-1" ALT="발행처">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBillToPartyNm" SIZE="20" MAXLENGTH="50" TAG="24" class = protected readonly = true TABINDEX="-1"></TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>계산서형태</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoTaxBillType" TAG="21X" VALUE="R" ID="rdoTaxBillType1">
									<LABEL FOR="rdoTaxBillType1">영수</LABEL>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoTaxBillType" TAG="21X" VALUE="D" CHECKED ID="rdoTaxBillType2">
									<LABEL FOR="rdoTaxBillType2">청구</LABEL>
								</TD>
								<TD CLASS=TD5 NOWRAP>발행여부</TD>
								<TD CLASS=TD6 NOWRAP>											
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoPostFlg" TAG="24X" class = protected readonly = true TABINDEX="-1" VALUE="Y" ID="rdoPostFlg1"> 
									<LABEL FOR="rdoPostFlg1">발행</LABEL>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoPostFlg" TAG="24X" class = protected readonly = true TABINDEX="-1" VALUE="N" CHECKED ID="rdoPostFlg2">
									<LABEL FOR="rdoPostFlg2">미발행</LABEL>
								</TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>발행일</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5311ma1_fpDateTime2_txtIssueDt.js'></script></TD>
								<TD CLASS=TD5 NOWRAP>공급가액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s5311ma1_fpDoubleSingle1_txtBillAmt.js'></script></TD>
											<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" class = protected readonly = true TABINDEX="-1" ALT="화폐"></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>
							<TR>	
								<TD CLASS=TD5 NOWRAP>세금신고사업장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTaxBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: left" TAG="23XXXU" class=required STYLE="text-transform:uppercase" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTaxBizArea" align=top TYPE="BUTTON" OnClick="VBScript:OpenTaxBizArea()">&nbsp;<INPUT TYPE=TEXT NAME="txtTaxBizAreaNm" SIZE=20 TAG="24" class = protected readonly = true TABINDEX="-1"></TD>
								<TD CLASS=TD5 NOWRAP>VAT액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s5311ma1_fpDoubleSingle1_txtVATAmt.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT적용기준</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATCalcType" TAG="21" VALUE="1"  ID="rdoVATCalcType1">
									<LABEL FOR="rdoVATCalcType1">개별</LABEL>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATCalcType" TAG="21" VALUE="2" CHECKED ID="rdoVATCalcType2">
									<LABEL FOR="rdoVATCalcType2">통합</LABEL>
								</TD>
								<TD CLASS=TD5 NOWRAP>공급가자국액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s5311ma1_fpDoubleSingle1_txtBillLocAmt.js'></script></TD>
											<TD>&nbsp;<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" class = protected readonly = true TABINDEX="-1" ALT="화폐"></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT포함구분</TD>
								<TD CLASS=TD6 NOWRAP>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATIncFlag" TAG="21" VALUE="1"  ID="rdoVATIncFlag1">
									<LABEL FOR="rdoVATIncFlag1">별도</LABEL>
									<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoVATIncFlag" TAG="21" VALUE="2" CHECKED ID="rdoVATIncFlag2">
									<LABEL FOR="rdoVATIncFlag2">포함</LABEL>
								</TD>
								<TD CLASS=TD5 NOWRAP>VAT자국액</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD><script language =javascript src='./js/s5311ma1_fpDoubleSingle1_txtVATLocAmt.js'></script></TD>
										</TR>
									</TABLE>
								</TD>				
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT유형</TD>
								<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="5" SIZE=10 ALT="VAT유형" tag="23XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBillHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;
											</TD>
											<TD>
												<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="25" SIZE=20 tag="24" class = protected readonly = true TABINDEX="-1">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="24XXXU" class = protected readonly = true TABINDEX="-1" class = protected readonly = true TABINDEX="-1" ALT="영업그룹">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24" class = protected readonly = true TABINDEX="-1"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>VAT율</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/s5311ma1_fpDoubleSingle4_txtVATRate.js'></script>&nbsp;%
								</TD>
								<TD CLASS=TD5 NOWRAP>비고</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRemark"  TYPE=TEXT MAXLENGTH=120 SIZE=42 TAG="21X" ALT="비고"></TD>
							</TR>
							<%Call SubFillRemBodyTD5656(10)%>
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
					<TD><BUTTON NAME="btnPosting" CLASS="CLSMBTN">발행</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck()">세금계산서내역등록</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtBillNoFlg" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHQueryMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHBillNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24"> 
<INPUT TYPE=HIDDEN NAME="txtHTaxBillNo" tag="24">

<!--추가 -->
<INPUT TYPE=HIDDEN NAME="txtMinor_cd" tag="24"> 
<INPUT TYPE=HIDDEN NAME="txtReference" tag="24"> 
<INPUT TYPE=HIDDEN NAME="txtVatCalcType" tag="24"> 
<INPUT TYPE=HIDDEN NAME="txtVatIncFlag" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

