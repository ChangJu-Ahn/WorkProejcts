<%@ LANGUAGE="VBSCRIPT" %>
<%'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5112MA2
'*  4. Program Name         : ���ܸ���ä�ǳ������ 
'*  5. Program Desc         :
'*  6. Comproxy List        : PS7G128.cSListBillDtlSvr,PS7G121.cSBillDtlSvr,PS7G115.cSPostOpenArSvr,PB3C104.cBLkUpItem
'*  7. Modified date(First) : 2002/11/14
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : AHN TAE HEE
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : 3rd ȭ�� Layout & ASP Coding
'*                            -2000/08/11 : 4th ȭ�� Layout
'*                            -2001/12/18 : Date ǥ������ 
'*                            -2001/12/26 : VAT �������� �߰� 
'*							  -2002/11/14 : UI���� ����	
'**********************************************************************************************%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="S5112ma2.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
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
									<td background="../../../CShared/../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���ܸ���ä�ǳ���</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								</TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenBillDtlRef">��������ä�ǳ�������</A></TD>
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
										<TD CLASS="TD5" NOWRAP>����ä�ǹ�ȣ</TD>
										<TD CLASS="TD6"><INPUT NAME="txtConBillNo" ALT="����ä�ǹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" STYLE="text-transform:uppercase" class=required><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSBillDtl" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBillDtl()"></TD>
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
									<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSoldToParty" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtSoldToPartyNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayTermsCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtPayTermsNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>ȭ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD><INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
												<TD>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtXchgRate" CLASS=FPDS100 tag="24X5" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>           
											</TR>
										</TABLE>  
									</TD>      
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVatType" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="24XXXU" class = protected readonly = true TABINDEX="-1">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>����ä�Ǳݾ�</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtOriginBillAmt" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								</TR>       
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT��</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtVatRate" CLASS=FPDS100 tag="24X5" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;%</TD>        
									<TD CLASS=TD5 NOWRAP>VAT�ݾ�</TD>
									<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtVatAmt" CLASS=FPDS140 tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoVatIncflag" id="rdoVatIncflag1" value="1" tag = "24">
										<label ID="lblVatIncFlag1" for="rdoVatIncflag1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoVatIncflag" id="rdoVatIncflag2" value="2" tag = "24" checked>
										<label ID="lblVatIncflag2" for="rdoVatIncflag2">����</label>
									</TD>
									<TD CLASS=TD5 NOWRAP>VAT�������</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoVatCalcType" id="rdoVatCalcType1" value="1" tag = "24">
										<label ID="lblVatCalcType1" for="rdoVatCalcType1">����</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoVatCalcType" id="rdoVatCalcType2" value="2" tag = "24" checked>
										<label ID="lblVatCalcType2" for="rdoVatCalcType2">����</label>
									</TD>
								</TR>
								<TR>
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" Title="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
						<TD><BUTTON NAME="btnPostFlag" CLASS="CLSMBTN">Ȯ��</BUTTON>&nbsp;
							<BUTTON NAME="btnGLView" CLASS="CLSMBTN">��ǥ��ȸ</BUTTON>&nbsp;
							<BUTTON NAME="btnPreRcptView" CLASS="CLSMBTN">��������Ȳ</BUTTON></TD>
						<TD WIDTH=* Align=Right><a href = "vbscript:JumpChgCheck(BIZ_BillHdr_JUMP_ID)">���ܸ���ä�ǵ��</a>&nbsp;|&nbsp;<a href = "vbscript:JumpChgCheck(BIZ_BillCollect_JUMP_ID)">����ä�Ǽ��ݳ������</a></TD>
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
	<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHBillNo" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtXchgOp" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtSts" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHRefFlag" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHBillType" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHBillTypeNm" tag="24" TABINDEX="-1">

	<INPUT TYPE=HIDDEN NAME="txtGLNo" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtTempGLNo" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtBatchNo" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHBillDt" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtRefBillNo" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtHPostFlag" tag="24" TABINDEX="-1">
	<P ID="divTextArea"></P>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
</DIV>

</BODY>
</HTML>
