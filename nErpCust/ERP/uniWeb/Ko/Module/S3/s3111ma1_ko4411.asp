<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : s3111ma1.asp	
'*  4. Program Name         : ���ֵ�� 
'*  5. Program Desc         : ���ֵ�� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/28
'*  8. Modified date(Last)  : 2005/11/25
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : NHG
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2005/11/25 -- 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="s3111ma1_ko441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '��: indicates that All variables must be declared in advance

'==========================================================================================================
Dim iDBSYSDate
Dim EndDate
iDBSYSDate = "<%=GetSvrDate%>"
'------ ��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'==========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����Ϲ�����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenSORef">��������</A>&nbsp;<A ID="txtOpenPrjRef" STYLE="DISPLAY: none" href="vbscript:OpenPrjRef">|&nbsp;������Ʈ����</A></TD>
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
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSo_no" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="12XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo frm1.txtConSo_no"></TD>
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
									<TD CLASS="TD5" NOWRAP>���ֹ�ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="���ֹ�ȣ" TYPE="Text" MAXLENGTH=18 SiZE=34 tag="25XXXU" STYLE="text-transform:uppercase" ></TD>
									<TD CLASS=TD5 NOWRAP>����Ȯ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoCfm_flag" id="rdoCfm_flag1" value="Y" tag = "24">
											<label for="rdoCfm_flag1">Ȯ��</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoCfm_flag" id="rdoCfm_flag2" value="N" tag = "24" checked>
											<label for="rdoCfm_flag2">��Ȯ��</label></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSo_Type" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 0" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtSo_TypeNm" TYPE="Text" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�ܰ�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPrice_flag" id="rdoPrice_flag1" value="Y" tag = "24" checked>
											<label for="rdoPrice_flag1">���ܰ�</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPrice_flag" id="rdoPrice_flag2" value="N" tag = "24">
											<label for="rdoPrice_flag2">���ܰ�</label>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtSo_dt" CLASS=FPDTYYYYMMDD tag="23X1" ALT="������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 NAME="txtReq_dlvy_dt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���ֹ���</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtCust_po_dt" CLASS=FPDTYYYYMMDD tag="21X1" ALT="���ֹ���" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>���ֹ���ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_po_no" TYPE="Text" MAXLENGTH="20" SIZE=20 ALT="���ֹ���ȣ" tag="25XXXU" STYLE="text-transform:uppercase"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�ֹ�ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSold_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase" ALT="�ֹ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp 0" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT NAME="txtSold_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Grp" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtSales_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtShip_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase" ALT="��ǰó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 0">&nbsp;<INPUT NAME="txtShip_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBill_to_party" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 1">&nbsp;<INPUT NAME="txtBill_to_partyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>����ó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPayer" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 11">&nbsp;<INPUT NAME="txtPayerNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>���ݱ׷�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTo_Biz_Grp" TYPE="Text" ALT="���ݱ׷�" MAXLENGTH="4" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 3">&nbsp;<INPUT NAME="txtTo_Biz_GrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ǹ�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeal_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase" ALT="�Ǹ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 3">&nbsp;<INPUT NAME="txtDeal_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>��۹��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTrans_Meth" ALT="��۹��" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 4">&nbsp;<INPUT NAME="txtTrans_Meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_terms" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU" class=required STYLE="text-transform:uppercase" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 5">&nbsp;<INPUT NAME="txtPay_terms_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�����Ⱓ</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtPay_dur" ALT="�����Ⱓ" style="HEIGHT: 20px; WIDTH: 150px" tag="21X6Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL>��</LABEL></TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT���Ա���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVat_Inc_Flag" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="VAT���Ա���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenOption 12">&nbsp;<INPUT NAME="txtVat_Inc_Flag_Nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>���ֱݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<INPUT NAME="txtDoc_cur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="23XXXU" class=required STYLE="text-transform:uppercase" ALT="ȭ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 2">&nbsp;
												</TD>
												<TD>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 NAME="txtNet_amt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVat_Type" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="21XXXU" STYLE="text-transform:uppercase" ALT="VAT����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 6">&nbsp;<INPUT NAME="txtVatTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>ȯ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtXchg_rate" style="HEIGHT: 20px; WIDTH: 150px" tag="22X5Z" ALT="ȯ��" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT��</TD>
									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 NAME="txtVat_rate" style="HEIGHT: 20px; WIDTH: 150px" tag="24X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>&nbsp;<LABEL><b>%</b></LABEL></TD>
									<TD CLASS=TD5 NOWRAP>�Ա�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_type" TYPE="TEXT" MAXLENGTH="5" SIZE=10 ALT="�Ա�����" TAG="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 10">&nbsp;<INPUT NAME="txtPay_Type_nm" TYPE="Text" MAXLENGTH="20" SIZE=24.5 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>VAT�ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 NAME="txtVat_amt" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>							
									</TD>
									<TD CLASS=TD5 NOWRAP>�����ڱ��ݾ�</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 NAME="txtNet_Amt_Loc" style="HEIGHT: 20px; WIDTH: 150px" tag="24X2Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>							
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ݰ�������</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txt_Payterms_txt" TYPE="Text" MAXLENGTH="120" SIZE=100 ALT="��ݰ�������" tag="21"></TD>
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>���</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="80" SIZE=100 ALT="���" tag="21"></TD>
								</TR>
							</TABLE>
							</DIV>
							
							<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIncoTerms" TYPE="Text" Alt="��������" MAXLENGTH="5" SIZE=10 tag="25XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRequried 2">&nbsp;<INPUT NAME="txtIncoTerms_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBeneficiary" TYPE="Text" ALT="������" MAXLENGTH="10" SIZE=10 tag="25XXXU" class=required STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBp 3">&nbsp;<INPUT NAME="txtBeneficiary_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 NAME="txtContract_dt" CLASS=FPDTYYYYMMDD tag="25X1" Alt="�����" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>��ȿ��</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime5 NAME="txtValid_dt" CLASS=FPDTYYYYMMDD tag="25X1" Alt="��ȿ��" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime6 NAME="txtship_dt" CLASS=FPDTYYYYMMDD tag="25X1" Alt="������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" TYPE="Text" ALT="������" MAXLENGTH="5" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinorCd 3">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>	
								<TR>
									<TD CLASS=TD5 NOWRAP>������������</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtShip_dt_txt" TYPE="Text" Alt="������������" MAXLENGTH="80" SIZE=100 tag="25"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoading_port_Cd" TYPE="Text" Alt="������" MAXLENGTH="5" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoading_port_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinorCd 2">&nbsp;<INPUT NAME="txtLoading_port_Nm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�۱�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSending_Bank" TYPE="Text" Alt="�۱�����" MAXLENGTH="10" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 7">&nbsp;<INPUT NAME="txtSending_Bank_nm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischge_port_Cd" TYPE="Text" Alt="������" MAXLENGTH="5" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischge_city_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMinorCd 1">&nbsp;<INPUT NAME="txtDischge_port_Nm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischge_city" TYPE="Text" Alt="��������" MAXLENGTH="30" SIZE=39.5 tag="25" STYLE="text-transform:uppercase"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPack_cond" TYPE="Text" Alt="��������" MAXLENGTH="5" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 8">&nbsp;<INPUT NAME="txtPack_cond_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>�˻���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInspect_meth" TYPE="Text" Alt="�˻���" MAXLENGTH="4" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenOption 9">&nbsp;<INPUT NAME="txtInspect_meth_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtManufacturer" TYPE="Text" ALT="������" MAXLENGTH="10" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBp 1">&nbsp;<INPUT NAME="txtManufacturer_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
									<TD CLASS=TD5 NOWRAP>������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAgent" TYPE="Text" ALT="������" MAXLENGTH="10" SIZE=10 tag="25XXXU" STYLE="text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBp 2">&nbsp;<INPUT NAME="txtAgent_nm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="24" class = protected readonly = true TABINDEX="-1"></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(11)%>
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
					<TD><BUTTON NAME="btnDNCheck" CLASS="CLSMBTN">���Ͽ�ûó��</BUTTON>&nbsp;
						<BUTTON NAME="btnConfirm" CLASS="CLSMBTN">Ȯ��ó��</BUTTON></TD>
					<TD WIDTH="*" ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">���ֳ������</A></TD>
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
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRadioDnParcel" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoSts" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHSONo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHDlvyLt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaintNo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="RdoConfirm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoTypeExportFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoTypeRetItemFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSoTypeCiFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtRetItemFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="RdoDnReq" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtProjectCd" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>	
</BODY>
</HTML>
